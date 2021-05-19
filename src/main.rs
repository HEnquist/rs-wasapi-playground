use std::mem;
use std::ptr;
use std::slice;
use std::fmt;
use wasapi::{
    PKEY_Device_FriendlyName,
    Windows::Win32::Media::Audio::CoreAudio::{
        eConsole, eRender, eCapture, IAudioClient, IAudioRenderClient, IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator, IMMDeviceCollection,
        AUDCLNT_SHAREMODE_EXCLUSIVE, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, DEVICE_STATE_ACTIVE, WAVE_FORMAT_EXTENSIBLE,
    },
    Windows::Win32::Media::Multimedia::{
        WAVEFORMATEX,
        WAVEFORMATEXTENSIBLE,
        WAVEFORMATEXTENSIBLE_0,
        WAVE_FORMAT_PCM,
        WAVE_FORMAT_IEEE_FLOAT,
        KSDATAFORMAT_SUBTYPE_PCM,
        KSDATAFORMAT_SUBTYPE_IEEE_FLOAT,
    },
    Windows::Win32::Storage::StructuredStorage::PROPVARIANT,
    Windows::Win32::Storage::StructuredStorage::STGM_READ,
    Windows::Win32::System::Com::{
        CLSIDFromProgID, CoInitializeEx, CoTaskMemAlloc, CoTaskMemFree, CLSCTX, CLSCTX_ALL,
        COINIT_MULTITHREADED,
    },
    Windows::Win32::System::PropertiesSystem::PropVariantToStringAlloc,
    Windows::Win32::System::PropertiesSystem::PROPERTYKEY,
    Windows::Win32::System::SystemServices::{PSTR, PWSTR},
    Windows::Win32::System::Threading::{
        CreateEventA,
        ResetEvent,
        SetEvent,
        WAIT_RETURN_CAUSE,
        WAIT_OBJECT_0,
        WaitForSingleObject,
        WaitForMultipleObjects,
    },
};
use widestring::U16CString;
use windows::Interface;
use std::error;

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

type Res<T> = Result<T, Box<dyn error::Error>>;

fn print_waveformat(wave_fmt: &WAVEFORMATEXTENSIBLE) {
    unsafe {
        println!("nAvgBytesPerSec {:?}", { wave_fmt.Format.nAvgBytesPerSec });
        println!("cbSize {:?}", { wave_fmt.Format.cbSize });
        println!("nBlockAlign {:?}", { wave_fmt.Format.nBlockAlign });
        println!("wBitsPerSample {:?}", { wave_fmt.Format.wBitsPerSample });
        println!("nSamplesPerSec {:?}", { wave_fmt.Format.nSamplesPerSec });
        println!("wFormatTag {:?}", { wave_fmt.Format.wFormatTag });
        println!("wValidBitsPerSample {:?}", { wave_fmt.Samples.wValidBitsPerSample });
        println!("SubFormat {:?}", { wave_fmt.SubFormat });
    }
}

fn make_waveformat(storebits: usize, validbits: usize, is_float: bool, samplerate: usize, channels: usize) -> WAVEFORMATEXTENSIBLE {
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
    WAVEFORMATEXTENSIBLE {
        Format: wave_format,
        Samples: sample,
        SubFormat: subformat,
        dwChannelMask: mask,
    }
}

fn get_iaudioclient(device: &IMMDevice) -> Res<IAudioClient> {
    let mut audio_client: mem::MaybeUninit<IAudioClient> = mem::MaybeUninit::zeroed();
    unsafe {
        device
            .Activate(
                &IAudioClient::IID,
                CLSCTX_ALL.0,
                ptr::null_mut(),
                audio_client.as_mut_ptr() as *mut _,
            )
            .ok()?;
        Ok(audio_client.assume_init())
    }
}

fn get_state(device: &IMMDevice) -> Res<u32> {
    let mut state: u32 = 0;
    unsafe  {
        device.GetState(&mut state).ok()?;
    }
    println!("state: {:?}", state);
    Ok(state)
}

fn get_friendlyname(device: &IMMDevice) -> Res<String> {
    let mut store = None;
    unsafe {
        device
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

fn get_id(device: &IMMDevice) -> Res<String> {
    let mut idstr = PWSTR::NULL;
    unsafe { 
        device.GetId(&mut idstr).ok()?;
    }
    let wide_id = unsafe { U16CString::from_ptr_str(idstr.0) };
    let id = wide_id.to_string_lossy();
    println!("id: {}", id);
    Ok(id)
}

fn is_supported_exclusive(audio_client: &IAudioClient, wave_fmt: &WAVEFORMATEXTENSIBLE) -> bool {
    let supported = unsafe { audio_client.IsFormatSupported(AUDCLNT_SHAREMODE_EXCLUSIVE, wave_fmt as *const _ as *const WAVEFORMATEX, ptr::null_mut()) };
    println!("supported {:?}\n", supported.ok());
    supported.ok().is_ok()
}

fn is_supported_shared(audio_client: &IAudioClient, wave_fmt: &WAVEFORMATEXTENSIBLE) -> Res<WAVEFORMATEXTENSIBLE> {
    let mut supported_format: mem::MaybeUninit<WAVEFORMATEXTENSIBLE> = mem::MaybeUninit::zeroed();
    unsafe { audio_client.IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, wave_fmt as *const _ as *const WAVEFORMATEX, &mut supported_format as *mut _ as *mut *mut WAVEFORMATEX).ok()? };
    let supported_format = unsafe {supported_format.assume_init()};
    Ok(supported_format)
}

fn get_periods(audio_client: &IAudioClient) -> Res<(i64, i64)> {
    let mut def_time = 0;
    let mut min_time = 0;
    unsafe { audio_client.GetDevicePeriod(&mut def_time, &mut min_time).ok()? };
    println!("default period {}, min period {}", def_time, min_time);
    Ok((def_time, min_time))
}

fn get_devices(capture: bool) -> Res<IMMDeviceCollection> {
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
    devs.ok_or(WasapiError::new("Failed to get devices").into())
}

fn get_device_with_name(devices: &IMMDeviceCollection, name: &str) -> Res<IMMDevice> {
    let mut count = 0;
    unsafe  { devices.GetCount(&mut count).ok()? };
    println!("nbr devices {}", count);
    for n in 0..count {
        let mut device = None;
        unsafe { devices.Item(n, &mut device).ok()? };
        let device = device.ok_or("Failed to get device")?;
        let devname = get_friendlyname(&device)?;
        if name == devname {
            return Ok(device)
        }
    }
    Err(WasapiError::new(format!("Unable to find device {}", name).as_str()).into())
}

fn main() -> Res<()> {
    unsafe {
        CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED).ok()?;
    }
    let mut device = None;
    let enumerator: IMMDeviceEnumerator = windows::create_instance(&MMDeviceEnumerator)?;
    unsafe {
        enumerator
            .GetDefaultAudioEndpoint(eRender, eConsole, &mut device)
            .ok()?;
        println!("{:?}", device);

        if let Some(device) = device {
            println!("{:?}", device);
            let mut store = None;
            device
                .OpenPropertyStore(STGM_READ as u32, &mut store)
                .ok()?;
            let mut state: u32 = 0;
            device.GetState(&mut state).ok()?;
            println!("{:?}", state);
            let mut prop: mem::MaybeUninit<PROPVARIANT> = mem::MaybeUninit::zeroed();
            let mut propstr = PWSTR::NULL;
            println!("read prop into {:?}", propstr);
            store
                .unwrap()
                .GetValue(&PKEY_Device_FriendlyName, prop.as_mut_ptr())
                .ok()?;
            let prop = prop.assume_init();
            println!("read prop");
            PropVariantToStringAlloc(&prop, &mut propstr).ok()?;
            // Get the buffer as a wide string
            let wide_string = U16CString::from_ptr_str(propstr.0);
            //let name = PWSTR::try_from(prop);
            println!("{}", wide_string.to_string_lossy());
        }
    }

    let devs = get_devices(false)?;
    let mut count = 0;
    unsafe  { devs.GetCount(&mut count).ok()? };
    println!("nbr devices {}", count);
    for n in 0..count {
        let mut device = None;
        unsafe { devs.Item(n, &mut device).ok()? };
        let device = device.ok_or("Failed to get device")?;
        let name = get_friendlyname(&device)?;
        let state = get_state(&device)?;
        let id = get_id(&device)?;

        let audio_client = get_iaudioclient(&device)?;

        let desired_format_ex = make_waveformat(16, 16, false, 48000, 2);
        let blockalign = {desired_format_ex.Format.nBlockAlign } as u32;
        print_waveformat(&desired_format_ex);

        let supported = is_supported_exclusive(&audio_client, &desired_format_ex);
        println!("supported {:?}\n", supported);

        let (def_time, min_time) = get_periods(&audio_client)?;
        println!("default period {}, min period {}", def_time, min_time);

        unsafe {
            audio_client.Initialize(AUDCLNT_SHAREMODE_EXCLUSIVE,
                AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                def_time,
                def_time,
                &desired_format_ex as *const _ as *const WAVEFORMATEX,
                std::ptr::null()).ok()?
        };

        let h_event = unsafe { CreateEventA(std::ptr::null_mut(), false, false, PSTR::default()) };

        unsafe { audio_client.SetEventHandle(h_event).ok()? };
            
        let mut bufferFrameCount = 0;
        unsafe { audio_client.GetBufferSize(&mut bufferFrameCount).ok()? };
        println!("bufferFrameCount {}",bufferFrameCount);

        let render_client: IAudioRenderClient = unsafe { audio_client.GetService()? };

        unsafe { audio_client.Start().ok()? };

        for n in 0..20 {
            let mut data = mem::MaybeUninit::uninit();
            unsafe { 
                render_client
                    .GetBuffer(bufferFrameCount, data.as_mut_ptr())
                    .ok()?
            };

            let mut dataptr = unsafe { data.assume_init() };
            let mut databuf = unsafe { slice::from_raw_parts_mut(dataptr, (bufferFrameCount*blockalign) as usize) };
            for m in 0..bufferFrameCount*blockalign/2 {
                databuf[m as usize] = 10;
            }
            unsafe { render_client.ReleaseBuffer(bufferFrameCount, 0) };
            println!("wrote frames");

            let retval = unsafe { WaitForSingleObject(h_event, 100) };
            if (retval != WAIT_OBJECT_0)
            {
                // Event handle timed out after a 2-second wait.
                unsafe { audio_client.Stop() };
                break;
            }

        }

        /*
        let mut desired_format_result: *mut WAVEFORMATEXTENSIBLE = ptr::null_mut();
        let supported = audio_client.IsFormatSupported(AUDCLNT_SHAREMODE_EXCLUSIVE, &desired_format_ex as *const _ as *const WAVEFORMATEX, &mut desired_format_result as *mut _ as *mut *mut WAVEFORMATEX);
        */

    }
    let device = get_device_with_name(&devs, "SPDIF Interface (FX-AUDIO-DAC-X6)")?;
    println!("done");
    Ok(())
}
