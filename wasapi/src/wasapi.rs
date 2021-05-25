use std::mem;
use std::ptr;
use std::slice;
use std::fmt;
use std::collections::VecDeque;
use widestring::U16CString;
use windows::Interface;
use std::error;
use crate::{
    PKEY_Device_FriendlyName,
    Windows::Win32::Media::Audio::CoreAudio::{
        eConsole, eRender, eCapture, IAudioClient, IAudioRenderClient, IAudioCaptureClient, IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator, IMMDeviceCollection,
        AUDCLNT_SHAREMODE_EXCLUSIVE, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, DEVICE_STATE_ACTIVE, WAVE_FORMAT_EXTENSIBLE,
    },
    Windows::Win32::Media::Multimedia::{
        WAVEFORMATEX,
        WAVEFORMATEXTENSIBLE,
        WAVEFORMATEXTENSIBLE_0,
        KSDATAFORMAT_SUBTYPE_PCM,
        KSDATAFORMAT_SUBTYPE_IEEE_FLOAT,
    },
    Windows::Win32::Storage::StructuredStorage::PROPVARIANT,
    Windows::Win32::Storage::StructuredStorage::STGM_READ,
    Windows::Win32::System::Com::CLSCTX_ALL,
    Windows::Win32::System::PropertiesSystem::PropVariantToStringAlloc,
    Windows::Win32::System::SystemServices::{PSTR, PWSTR, HANDLE, S_OK, S_FALSE},
    Windows::Win32::System::Threading::{
        CreateEventA,
        WAIT_OBJECT_0,
        WaitForSingleObject,
    },
};

type WasapiRes<T> = Result<T, Box<dyn error::Error>>;

// Error returned by the Wasapi crate.
#[derive(Debug)]
pub struct WasapiError {
    desc: String,
}

impl fmt::Display for WasapiError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.desc)
    }
}

impl error::Error for WasapiError {
    fn description(&self) -> &str {
        &self.desc
    }
}

impl WasapiError {
    pub fn new(desc: &str) -> Self {
        WasapiError {
            desc: desc.to_owned(),
        }
    }
}

// Audio direction, playback or capture.
#[derive(Clone)]
pub enum Direction {
    Render,
    Capture,
}

// Sharemode for device
#[derive(Clone)]
pub enum ShareMode {
    Shared,
    Exclusive,
}

// Sample type, float or integer
#[derive(Clone)]
pub enum SampleType {
    Float,
    Int,
}

// Get the default playback or capture device
pub fn get_default_device(direction: &Direction) -> WasapiRes<Device> {
    let dir = match direction {
        Direction::Capture => eCapture,
        Direction::Render => eRender,
    };
    let mut device = None;
    let enumerator: IMMDeviceEnumerator = windows::create_instance(&MMDeviceEnumerator)?;
    unsafe {
        enumerator
            .GetDefaultAudioEndpoint(dir, eConsole, &mut device)
            .ok()?;
        println!("{:?}", device);
    }
    match device {
        Some(dev) => Ok(Device{ device: dev, direction: direction.clone()}),
        None => Err(WasapiError::new("Failed to get default device").into()),
    }
}

// Struct wrapping an IMMDeviceCollection.
pub struct DeviceCollection {
    collection: IMMDeviceCollection,
    direction: Direction,
}

impl DeviceCollection {
    // Get an IMMDeviceCollection of all active playback or capture devices
    pub fn new(direction: &Direction) -> WasapiRes<DeviceCollection> {
        let dir = match direction {
            Direction::Capture => eCapture,
            Direction::Render => eRender,
        };
        let enumerator: IMMDeviceEnumerator = windows::create_instance(&MMDeviceEnumerator)?;
        let mut devs = None;
        unsafe {
            enumerator
                .EnumAudioEndpoints(dir, DEVICE_STATE_ACTIVE, &mut devs)
                .ok()?;
        }
        match devs {
            Some(collection) => Ok(DeviceCollection { collection, direction: direction.clone()}),
            None => Err(WasapiError::new("Failed to get devices").into()),
        }
    }

    // Get the number of devices in an IMMDeviceCollection
    pub fn get_nbr_devices(&self) -> WasapiRes<u32> {
        let mut count = 0;
        unsafe  { self.collection.GetCount(&mut count).ok()? };
        Ok(count)
    }

    // Get a device from an IMMDeviceCollection using index
    pub fn get_device_at_index(&self, idx: u32) -> WasapiRes<Device> {
        let mut dev = None;
        unsafe { self.collection.Item(idx, &mut dev).ok()? };
        match dev {
            Some(device) => Ok(Device { device, direction: self.direction.clone()}),
            None => Err(WasapiError::new("Failed to get device").into()),
        }
    }

    // Get a device from an IMMDeviceCollection using name
    pub fn get_device_with_name(&self, name: &str) -> WasapiRes<Device> {
        let mut count = 0;
        unsafe  { self.collection.GetCount(&mut count).ok()? };
        println!("nbr devices {}", count);
        for n in 0..count {
            let device = self.get_device_at_index(n)?;
            let devname = device.get_friendlyname()?;
            if name == devname {
                return Ok(device)
            }
        }
        Err(WasapiError::new(format!("Unable to find device {}", name).as_str()).into())
    }
}

// Struct wrapping an IMMDevice.
pub struct Device {
    device: IMMDevice,
    direction: Direction,
}

impl Device {
    // Get an IAudioClient from an IMMDevice
    pub fn get_iaudioclient(&self) -> WasapiRes<AudioClient> {
        let mut audio_client: mem::MaybeUninit<IAudioClient> = mem::MaybeUninit::zeroed();
        unsafe {
            self.device
                .Activate(
                    &IAudioClient::IID,
                    CLSCTX_ALL.0,
                    ptr::null_mut(),
                    audio_client.as_mut_ptr() as *mut _,
                )
                .ok()?;
            Ok(AudioClient { client: audio_client.assume_init(), direction: self.direction.clone(), sharemode: None})
        }
    }

    // Read state from an IMMDevice
    pub fn get_state(&self) -> WasapiRes<u32> {
        let mut state: u32 = 0;
        unsafe  {
            self.device.GetState(&mut state).ok()?;
        }
        println!("state: {:?}", state);
        Ok(state)
    }

    // Read the FrienlyName of an IMMDevice 
    pub fn get_friendlyname(&self) -> WasapiRes<String> {
        let mut store = None;
        unsafe {
            self.device
                .OpenPropertyStore(STGM_READ as u32, &mut store)
                .ok()?;
        }
        let mut prop: mem::MaybeUninit<PROPVARIANT> = mem::MaybeUninit::zeroed();
        let mut propstr = PWSTR::NULL;
        let store = store.ok_or("Failed to get store")?;
        unsafe {
            store
                .GetValue(&PKEY_Device_FriendlyName, prop.as_mut_ptr())
                .ok()?;
            let prop = prop.assume_init();
            PropVariantToStringAlloc(&prop, &mut propstr).ok()?;
        }
        let wide_name = unsafe { U16CString::from_ptr_str(propstr.0) };
        let name =  wide_name.to_string_lossy();
        println!("name: {}", name);
        Ok(name)
    }

    // Get the Id of an IMMDevice
    pub fn get_id(&self) -> WasapiRes<String> {
        let mut idstr = PWSTR::NULL;
        unsafe { 
            self.device.GetId(&mut idstr).ok()?;
        }
        let wide_id = unsafe { U16CString::from_ptr_str(idstr.0) };
        let id = wide_id.to_string_lossy();
        println!("id: {}", id);
        Ok(id)
    }
}

// Return type for is_supported.
pub enum FormatSupported {
    // The format is supported as it is
    Yes,
    // The format is not supported as it is, this is the closest match.
    ClosestMatch(WaveFormat),
}

// Struct wrapping an IAudioClient.
pub struct AudioClient {
    client: IAudioClient,
    direction: Direction,
    sharemode: Option<ShareMode>,
}

impl AudioClient {

    // Check if a format is supported, return the nearest match (identical to requested for exclusive mode)
    pub fn is_supported(&self, wave_fmt: &WaveFormat, sharemode: &ShareMode) -> WasapiRes<FormatSupported> {
        let supported = match sharemode {
            ShareMode::Exclusive => {
                unsafe { self.client.IsFormatSupported(AUDCLNT_SHAREMODE_EXCLUSIVE, wave_fmt.as_waveformatex_ptr(), ptr::null_mut()).ok()? };
                FormatSupported::Yes
            },
            ShareMode::Shared => {
                let mut supported_format: mem::MaybeUninit<*mut WAVEFORMATEXTENSIBLE> = mem::MaybeUninit::zeroed();
                let res = unsafe { self.client.IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, wave_fmt.as_waveformatex_ptr(), &mut supported_format as *mut _ as *mut *mut WAVEFORMATEX) };
                res.ok()?;
                if res == S_OK {
                    println!("supported");
                    FormatSupported::Yes
                }
                else if res == S_FALSE {
                    println!("not supported");
                    let new_fmt = unsafe {supported_format.assume_init().read()};
                    FormatSupported::ClosestMatch(WaveFormat{ wave_fmt: new_fmt })
                }
                else {
                    return Err(WasapiError::new("Unsupported format").into());
                }
            }
        };
        Ok(supported)
    }

    // Get default and minimum periods in 100-nanosecond units
    pub fn get_periods(&self) -> WasapiRes<(i64, i64)> {
        let mut def_time = 0;
        let mut min_time = 0;
        unsafe { self.client.GetDevicePeriod(&mut def_time, &mut min_time).ok()? };
        println!("default period {}, min period {}", def_time, min_time);
        Ok((def_time, min_time))
    }



    // Initialize an IAudioClient
    pub fn initialize_client(&mut self, wavefmt: &WaveFormat, period: i64, direction: &Direction, sharemode: &ShareMode) -> WasapiRes<()> {
        let streamflags = match (&self.direction, direction, sharemode) {
            (Direction::Render, Direction::Capture, ShareMode::Shared) => AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_LOOPBACK,
            (Direction::Render, Direction::Capture, ShareMode::Exclusive) => {
                return Err(WasapiError::new("Cant use Loopback with exclusive mode").into());
            },
            (Direction::Capture, Direction::Render, _) => {
                return Err(WasapiError::new("Cant render to a capture device").into());
            },
            _ => AUDCLNT_STREAMFLAGS_EVENTCALLBACK
        };
        let mode = match sharemode {
            ShareMode::Exclusive => AUDCLNT_SHAREMODE_EXCLUSIVE,
            ShareMode::Shared => AUDCLNT_SHAREMODE_SHARED,
        };
        self.sharemode = Some(sharemode.clone());
        unsafe {
            self.client.Initialize(mode,
                streamflags,
                period,
                period,
                wavefmt.as_waveformatex_ptr(),
                std::ptr::null()).ok()?;
        }
        Ok(())
    }

    // Create an return an event handle for an IAudioClient
    pub fn set_get_eventhandle(&self) -> WasapiRes<Handle> {
        let h_event = unsafe { CreateEventA(std::ptr::null_mut(), false, false, PSTR::default()) };
        unsafe { self.client.SetEventHandle(h_event).ok()? };
        Ok(Handle {handle: h_event})
    }

    // Get buffer size in frames
    pub fn get_bufferframecount(&self) -> WasapiRes<u32> {
        let mut buffer_frame_count = 0;
        unsafe { self.client.GetBufferSize(&mut buffer_frame_count).ok()? };
        println!("buffer_frame_count {}",buffer_frame_count);
        Ok(buffer_frame_count)
    }

    // Get current padding in frames
    pub fn get_current_padding(&self) -> WasapiRes<u32> {
        let mut padding_count = 0;
        unsafe { self.client.GetCurrentPadding(&mut padding_count).ok()? };
        println!("padding_count {}",padding_count);
        Ok(padding_count)
    }

    // Get buffer size minus padding in frames
    pub fn get_available_frames(&self) -> WasapiRes<u32> {
        let frames = match self.sharemode {
            Some(ShareMode::Exclusive) => {
                let mut buffer_frame_count = 0;
                unsafe { self.client.GetBufferSize(&mut buffer_frame_count).ok()? };
                println!("buffer_frame_count {}",buffer_frame_count);
                buffer_frame_count
            },
            Some(ShareMode::Shared) => {
                let mut padding_count = 0;
                let mut buffer_frame_count = 0;
                unsafe { self.client.GetBufferSize(&mut buffer_frame_count).ok()? };
                unsafe { self.client.GetCurrentPadding(&mut padding_count).ok()? };
                buffer_frame_count-padding_count
            },
            _ => return Err(WasapiError::new("Client has not been initialized").into()),
        };
        Ok(frames)
    }


    // Start the stream on an IAudioClient
    pub fn start_stream(&self) -> WasapiRes<()> {
        unsafe { self.client.Start().ok()? };
        Ok(())
    }

    // Stop the stream on an IAudioClient
    pub fn stop_stream(&self) -> WasapiRes<()> {
        unsafe { self.client.Stop().ok()? };
        Ok(())
    }

    pub fn get_audiorenderclient(&self) -> WasapiRes<AudioRenderClient> {
        let renderclient: Option<IAudioRenderClient> = unsafe { self.client.GetService().ok() };
        match renderclient {
            Some(client) => Ok(AudioRenderClient {client}),
            None => Err(WasapiError::new("Failed getting IAudioRenderClient").into()),
        }
    }

    pub fn get_audiocaptureclient(&self) -> WasapiRes<AudioCaptureClient> {
        let renderclient: Option<IAudioCaptureClient> = unsafe { self.client.GetService().ok() };
        match renderclient {
            Some(client) => Ok(AudioCaptureClient {client}),
            None => Err(WasapiError::new("Failed getting IAudioCaptureClient").into()),
        }
    }
}

// Struct wrapping an IAudioRenderClient.
pub struct AudioRenderClient {
    client: IAudioRenderClient,
}

impl AudioRenderClient {
    // Write raw bytes data to a device from a slice
    pub fn write_to_device(&self, nbr_frames: usize, byte_per_frame: usize, data: &[u8]) -> WasapiRes<()> {
        let nbr_bytes = nbr_frames * byte_per_frame;
        if nbr_bytes != data.len() {
            return Err(WasapiError::new(format!("Wrong length of data, got {}, expected {}", data.len(), nbr_bytes).as_str()).into());
        }
        let mut buffer = mem::MaybeUninit::uninit();
        unsafe { 
            self.client
                .GetBuffer(nbr_frames as u32, buffer.as_mut_ptr())
                .ok()?
        };
        let bufferptr = unsafe { buffer.assume_init() };
        let bufferslice = unsafe { slice::from_raw_parts_mut(bufferptr, nbr_bytes) };
        bufferslice.copy_from_slice(data);
        unsafe { self.client.ReleaseBuffer(nbr_frames as u32, 0).ok()? };
        println!("wrote frames");
        Ok(())
    }

    // Write raw bytes data to a device from a deque
    pub fn write_to_device_from_deque(&self, nbr_frames: usize, byte_per_frame: usize, data: &mut VecDeque<u8>) -> WasapiRes<()> {
        let nbr_bytes = nbr_frames * byte_per_frame;
        if nbr_bytes > data.len() {
            return Err(WasapiError::new(format!("To little data, got {}, need {}", data.len(), nbr_bytes).as_str()).into());
        }
        let mut buffer = mem::MaybeUninit::uninit();
        println!("get buffer");
        unsafe { 
            let val = self.client
                .GetBuffer(nbr_frames as u32, buffer.as_mut_ptr());
            println!("res {:?}", val);
                
            val.ok()?
        };
        println!("copy to buffer");
        let bufferptr = unsafe { buffer.assume_init() };
        let bufferslice = unsafe { slice::from_raw_parts_mut(bufferptr, nbr_bytes) };
        for element in bufferslice.iter_mut() {
            *element = data.pop_front().unwrap();
        }
        unsafe { self.client.ReleaseBuffer(nbr_frames as u32, 0).ok()? };
        println!("wrote frames");
        Ok(())
    }
}

// Struct wrapping an IAudioCaptureClient.
pub struct AudioCaptureClient {
    client: IAudioCaptureClient,
}

impl AudioCaptureClient {
    // Get number of frames in next packet, only works in shared mode
    pub fn get_next_nbr_frames(&self) -> WasapiRes<u32> {
        let mut nbr_frames = 0;
        unsafe {self.client.GetNextPacketSize(&mut nbr_frames).ok()?};
        Ok(nbr_frames)
    }

    // Read raw bytes data from a device into a slice
    pub fn read_from_device(&self, bytes_per_frame: usize, data: &mut [u8]) -> WasapiRes<()> {
        let data_len_in_frames = data.len() / bytes_per_frame;
        let mut buffer = mem::MaybeUninit::uninit();
        let mut nbr_frames_returned = 0;
        unsafe { 
            self.client
                .GetBuffer(buffer.as_mut_ptr(), &mut nbr_frames_returned, &mut 0, ptr::null_mut(), ptr::null_mut())
                .ok()?
        };
        if data_len_in_frames != nbr_frames_returned as usize {
            return Err(WasapiError::new(format!("Wrong length of data, got {} frames, expected {} frames", data_len_in_frames, nbr_frames_returned).as_str()).into());
        }
        let len_in_bytes = nbr_frames_returned as usize * bytes_per_frame;
        let bufferptr = unsafe { buffer.assume_init() };
        let bufferslice = unsafe { slice::from_raw_parts(bufferptr, len_in_bytes) };
        data.copy_from_slice(bufferslice);
        unsafe { self.client.ReleaseBuffer(nbr_frames_returned).ok()? };
        println!("wrote frames");
        Ok(())
    }

    // Write raw bytes data to a device from a deque
    pub fn read_from_device_to_deque(&self, bytes_per_frame: usize, data: &mut VecDeque<u8>) -> WasapiRes<()> {
        let mut buffer = mem::MaybeUninit::uninit();
        let mut nbr_frames_returned = 0;
        unsafe { 
            self.client
                .GetBuffer(buffer.as_mut_ptr(), &mut nbr_frames_returned, &mut 0, ptr::null_mut(), ptr::null_mut())
                .ok()?
        };
        let len_in_bytes = nbr_frames_returned as usize * bytes_per_frame;
        let bufferptr = unsafe { buffer.assume_init() };
        let bufferslice = unsafe { slice::from_raw_parts(bufferptr, len_in_bytes) };
        for element in bufferslice.iter() {
            data.push_back(*element);
        }
        unsafe { self.client.ReleaseBuffer(nbr_frames_returned).ok()? };
        //println!("wrote frames");
        Ok(())
    }
}

// Struct wrapping a HANDLE (event handle).
pub struct Handle {
    handle: HANDLE,
}

impl Handle {
    // Wait for an event on a handle
    pub fn wait_for_event(&self, timeout_ms: u32) -> WasapiRes<()> {
        let retval = unsafe { WaitForSingleObject(self.handle, timeout_ms) };
        if retval != WAIT_OBJECT_0
        {
            return Err(WasapiError::new(format!("Wait timed out").as_str()).into());
        }
        Ok(())
    }
}

// Struct wrapping a WAVEFORMATEXTENSIBLE format descriptor.
#[derive(Clone)]
pub struct WaveFormat {
    wave_fmt: WAVEFORMATEXTENSIBLE,
}

impl WaveFormat {
    // Print all fields, for debugging
    pub fn print_waveformat(&self) {
        unsafe {
            println!("nAvgBytesPerSec {:?}", { self.wave_fmt.Format.nAvgBytesPerSec });
            println!("cbSize {:?}", { self.wave_fmt.Format.cbSize });
            println!("nBlockAlign {:?}", { self.wave_fmt.Format.nBlockAlign });
            println!("wBitsPerSample {:?}", { self.wave_fmt.Format.wBitsPerSample });
            println!("nSamplesPerSec {:?}", { self.wave_fmt.Format.nSamplesPerSec });
            println!("wFormatTag {:?}", { self.wave_fmt.Format.wFormatTag });
            println!("wValidBitsPerSample {:?}", { self.wave_fmt.Samples.wValidBitsPerSample });
            println!("SubFormat {:?}", { self.wave_fmt.SubFormat });
            println!("nChannels {:?}", { self.wave_fmt.Format.nChannels });
            println!("dwChannelMask {:?}", { self.wave_fmt.dwChannelMask });
        }
    }

    // Build a WAVEFORMATEXTENSIBLE struct for the given parameters
    pub fn new(storebits: usize, validbits: usize, sample_type: &SampleType, samplerate: usize, channels: usize) -> Self {
        let blockalign = channels * storebits / 8;
        let byterate = samplerate * blockalign;

        let wave_format = WAVEFORMATEX {
            cbSize: 22,
            nAvgBytesPerSec: byterate as u32,
            nBlockAlign: blockalign as u16, 
            nChannels: channels as u16,
            nSamplesPerSec: samplerate as u32,
            wBitsPerSample: storebits as u16,
            wFormatTag: WAVE_FORMAT_EXTENSIBLE as u16,
        };
        let sample = WAVEFORMATEXTENSIBLE_0 {
            wValidBitsPerSample: validbits as u16,
        };
        let subformat = match sample_type {
            SampleType::Float => KSDATAFORMAT_SUBTYPE_IEEE_FLOAT,
            SampleType::Int => KSDATAFORMAT_SUBTYPE_PCM,
        };
        let mut mask = 0;
        for n in 0..channels {
            mask += 1<<n;
        }
        let wave_fmt = WAVEFORMATEXTENSIBLE {
            Format: wave_format,
            Samples: sample,
            SubFormat: subformat,
            dwChannelMask: mask,
        };
        WaveFormat{ wave_fmt }
    }

    // get a pointer of type WAVEFORMATEX, used internally
    pub fn as_waveformatex_ptr(&self) -> *const WAVEFORMATEX {
        &self.wave_fmt as *const _ as *const WAVEFORMATEX
    }

    // Read nBlockAlign.
    pub fn get_blockalign(&self) -> u32 {
        self.wave_fmt.Format.nBlockAlign as u32
    }

    // Read nAvgBytesPerSec.
    pub fn get_avgbytespersec(&self) -> u32 {
        self.wave_fmt.Format.nAvgBytesPerSec
    }

    // Read wBitsPerSample.
    pub fn get_bitspersample(&self) -> u16 {
        self.wave_fmt.Format.wBitsPerSample
    }

    // Read wValidBitsPerSample.
    pub fn get_validbitspersample(&self) -> u16 {
        unsafe { self.wave_fmt.Samples.wValidBitsPerSample }
    }

    // Read nSamplesPerSec.
    pub fn get_samplespersec(&self) -> u32 {
        self.wave_fmt.Format.nSamplesPerSec
    }

    // Read nChannels.
    pub fn get_nchannels(&self) -> u16 {
        self.wave_fmt.Format.nChannels
    }

    // Read dwChannelMask.
    pub fn get_dwchannelmask(&self) -> u32 {
        self.wave_fmt.dwChannelMask
    }

    // Read SubFormat.
    pub fn get_subformat(&self) -> WasapiRes<SampleType> {
        let subfmt = match self.wave_fmt.SubFormat {
            KSDATAFORMAT_SUBTYPE_IEEE_FLOAT  => SampleType::Float,
            KSDATAFORMAT_SUBTYPE_PCM => SampleType::Int,
            _ => {
                return Err(WasapiError::new(format!("Unknown subformat {:?}", {self.wave_fmt.SubFormat}).as_str()).into());
            },
        };
        Ok(subfmt)
    }
}
