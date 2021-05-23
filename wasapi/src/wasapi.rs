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
        eRender, eCapture, IAudioClient, IAudioRenderClient, IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator, IMMDeviceCollection,
        AUDCLNT_SHAREMODE_EXCLUSIVE, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, DEVICE_STATE_ACTIVE, WAVE_FORMAT_EXTENSIBLE,
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
    Windows::Win32::System::SystemServices::{PSTR, PWSTR, HANDLE},
    Windows::Win32::System::Threading::{
        CreateEventA,
        WAIT_OBJECT_0,
        WaitForSingleObject,
    },
};

type Res<T> = Result<T, Box<dyn error::Error>>;

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

pub struct DeviceCollection {
    collection: IMMDeviceCollection,
}

impl DeviceCollection {
    // Get an IMMDeviceCollection of all active playback or capture devices
    pub fn new(capture: bool) -> Res<DeviceCollection> {
        let direction = if capture {
            eCapture
        }
        else {
            eRender
        };
        let enumerator: IMMDeviceEnumerator = windows::create_instance(&MMDeviceEnumerator)?;
        let mut devs = None;
        unsafe {
            enumerator
                .EnumAudioEndpoints(direction, DEVICE_STATE_ACTIVE, &mut devs)
                .ok()?;
        }
        match devs {
            Some(collection) => Ok(DeviceCollection { collection}),
            None => Err(WasapiError::new("Failed to get devices").into()),
        }
    }

    // Get the number of devices in an IMMDeviceCollection
    pub fn get_nbr_devices(&self) -> Res<u32> {
        let mut count = 0;
        unsafe  { self.collection.GetCount(&mut count).ok()? };
        Ok(count)
    }

    // Get a device from an IMMDeviceCollection using index
    pub fn get_device_at_index(&self, idx: u32) -> Res<Device> {
        let mut dev = None;
        unsafe { self.collection.Item(idx, &mut dev).ok()? };
        match dev {
            Some(device) => Ok(Device { device}),
            None => Err(WasapiError::new("Failed to get device").into()),
        }
    }

    // Get a device from an IMMDeviceCollection using name
    pub fn get_device_with_name(&self, name: &str) -> Res<Device> {
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

pub struct Device {
    device: IMMDevice,
}

impl Device {
    // Get an IAudioClient from an IMMDevice
    pub fn get_iaudioclient(&self) -> Res<AudioClient> {
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
            Ok(AudioClient { client: audio_client.assume_init()})
        }
    }

    // Read state from an IMMDevice
    pub fn get_state(&self) -> Res<u32> {
        let mut state: u32 = 0;
        unsafe  {
            self.device.GetState(&mut state).ok()?;
        }
        println!("state: {:?}", state);
        Ok(state)
    }

    // Read the FrienlyName of an IMMDevice 
    pub fn get_friendlyname(&self) -> Res<String> {
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
    pub fn get_id(&self) -> Res<String> {
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


pub struct AudioClient {
    client: IAudioClient,
}

impl AudioClient {
    // Check if a format is supported in exclusive mode
    pub fn is_supported_exclusive(&self, wave_fmt: &WaveFormat) -> bool {
        let supported = unsafe { self.client.IsFormatSupported(AUDCLNT_SHAREMODE_EXCLUSIVE, wave_fmt.as_waveformatex_ptr(), ptr::null_mut()) };
        println!("supported {:?}\n", supported.ok());
        supported.ok().is_ok()
    }

    // Get the nearest supported format in shared mode
    pub fn is_supported_shared(&self, wave_fmt: &WaveFormat) -> Res<WaveFormat> {
        let mut supported_format: mem::MaybeUninit<WAVEFORMATEXTENSIBLE> = mem::MaybeUninit::zeroed();
        unsafe { self.client.IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, wave_fmt.as_waveformatex_ptr(), &mut supported_format as *mut _ as *mut *mut WAVEFORMATEX).ok()? };
        let supported_format = unsafe {supported_format.assume_init()};
        Ok(WaveFormat{ wave_fmt: supported_format})
    }

    // Get default and minimum periods in 100-nanosecond units
    pub fn get_periods(&self) -> Res<(i64, i64)> {
        let mut def_time = 0;
        let mut min_time = 0;
        unsafe { self.client.GetDevicePeriod(&mut def_time, &mut min_time).ok()? };
        println!("default period {}, min period {}", def_time, min_time);
        Ok((def_time, min_time))
    }



    // Initialize an IAudioClient
    pub fn initialize_client(&self, wavefmt: &WaveFormat, period: i64) -> Res<()> {
        unsafe {
            self.client.Initialize(AUDCLNT_SHAREMODE_EXCLUSIVE,
                AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                period,
                period,
                wavefmt.as_waveformatex_ptr(),
                std::ptr::null()).ok()?;
        }
        Ok(())
    }

    // Create an return an event handle for an IAudioClient
    pub fn set_get_eventhandle(&self) -> Res<Handle> {
        let h_event = unsafe { CreateEventA(std::ptr::null_mut(), false, false, PSTR::default()) };
        unsafe { self.client.SetEventHandle(h_event).ok()? };
        Ok(Handle {handle: h_event})
    }

    // Get buffer size in frames
    pub fn get_bufferframecount(&self) -> Res<u32> {
        let mut buffer_frame_count = 0;
        unsafe { self.client.GetBufferSize(&mut buffer_frame_count).ok()? };
        println!("buffer_frame_count {}",buffer_frame_count);
        Ok(buffer_frame_count)
    }

    // Start the stream on an IAudioClient
    pub fn start_stream(&self) -> Res<()> {
        unsafe { self.client.Start().ok()? };
        Ok(())
    }

    // Stop the stream on an IAudioClient
    pub fn stop_stream(&self) -> Res<()> {
        unsafe { self.client.Stop().ok()? };
        Ok(())
    }

    pub fn get_audiorenderclient(&self) -> Res<AudioRenderClient> {
        let renderclient: Option<IAudioRenderClient> = unsafe { self.client.GetService().ok() };
        match renderclient {
            Some(client) => Ok(AudioRenderClient {client}),
            None => Err(WasapiError::new("Failed getting IAudioRenderClient").into()),
        }
    }
}

pub struct AudioRenderClient {
    client: IAudioRenderClient,
}

impl AudioRenderClient {
    // Write raw bytes data to a device from a slice
    pub fn write_to_device(&self, nbr_frames: usize, byte_per_frame: usize, data: &[u8]) -> Res<()> {
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
    pub fn write_to_device_from_deque(&self, nbr_frames: usize, byte_per_frame: usize, data: &mut VecDeque<u8>) -> Res<()> {
        let nbr_bytes = nbr_frames * byte_per_frame;
        if nbr_bytes > data.len() {
            return Err(WasapiError::new(format!("To little data, got {}, need {}", data.len(), nbr_bytes).as_str()).into());
        }
        let mut buffer = mem::MaybeUninit::uninit();
        unsafe { 
            self.client
                .GetBuffer(nbr_frames as u32, buffer.as_mut_ptr())
                .ok()?
        };
        let bufferptr = unsafe { buffer.assume_init() };
        let bufferslice = unsafe { slice::from_raw_parts_mut(bufferptr, nbr_bytes) };
        for element in bufferslice.iter_mut() {
            *element = data.pop_front().unwrap();
        }
        unsafe { self.client.ReleaseBuffer(nbr_frames as u32, 0).ok()? };
        //println!("wrote frames");
        Ok(())
    }
}

pub struct Handle {
    handle: HANDLE,
}

impl Handle {
    // Wait for an event on a handle
    pub fn wait_for_event(&self, timeout_ms: u32) -> Res<()> {
        let retval = unsafe { WaitForSingleObject(self.handle, timeout_ms) };
        if retval != WAIT_OBJECT_0
        {
            return Err(WasapiError::new(format!("Wait timed out").as_str()).into());
        }
        Ok(())
    }
}

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
        }
    }

    // Build a WAVEFORMATEXTENSIBLE struct for the given parameters
    pub fn new(storebits: usize, validbits: usize, is_float: bool, samplerate: usize, channels: usize) -> Self {
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
        let subformat = if is_float {
            KSDATAFORMAT_SUBTYPE_IEEE_FLOAT
        }
        else {
            KSDATAFORMAT_SUBTYPE_PCM
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

    pub fn as_waveformatex_ptr(&self) -> *const WAVEFORMATEX {
        &self.wave_fmt as *const _ as *const WAVEFORMATEX
    }

    pub fn get_blockalign(&self) -> u32 {
        self.wave_fmt.Format.nBlockAlign as u32
    }
}

