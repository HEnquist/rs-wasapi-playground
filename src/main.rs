use std::mem;
use std::ptr;
use wasapi::{
    PKEY_Device_FriendlyName,
    Windows::Win32::Media::Audio::CoreAudio::{
        eConsole, eRender, IAudioClient, IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator,
        AUDCLNT_SHAREMODE_EXCLUSIVE, AUDCLNT_SHAREMODE_SHARED, DEVICE_STATE_ACTIVE,
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
    Windows::Win32::System::SystemServices::PWSTR,
};
use widestring::U16CString;
use windows::Interface;

fn main() -> windows::Result<()> {
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
    unsafe {
        let mut devs = None;
        enumerator
            .EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &mut devs)
            .ok()?;
        let devs = devs.unwrap();
        let mut count = 0;
        devs.GetCount(&mut count).ok()?;
        println!("nbr devices {}", count);
        for n in 0..count {
            let mut device = None;
            devs.Item(n, &mut device).unwrap();
            let device = device.unwrap();
            let mut idstr = PWSTR::NULL;
            device.GetId(&mut idstr).ok()?;
            let wide_id = U16CString::from_ptr_str(idstr.0);
            println!("id: {}", wide_id.to_string_lossy());
            let mut store = None;
            device
                .OpenPropertyStore(STGM_READ as u32, &mut store)
                .ok()?;
            let mut state: u32 = 0;
            device.GetState(&mut state).ok()?;
            println!("state: {:?}", state);
            let mut prop: mem::MaybeUninit<PROPVARIANT> = mem::MaybeUninit::zeroed();
            let mut propstr = PWSTR::NULL;
            store
                .unwrap()
                .GetValue(&PKEY_Device_FriendlyName, prop.as_mut_ptr())
                .ok()?;
            let prop = prop.assume_init();
            PropVariantToStringAlloc(&prop, &mut propstr).ok()?;
            let wide_name = U16CString::from_ptr_str(propstr.0);
            println!("name: {}", wide_name.to_string_lossy());

            let mut audio_client: mem::MaybeUninit<IAudioClient> = mem::MaybeUninit::zeroed();

            device
                .Activate(
                    &IAudioClient::IID,
                    CLSCTX_ALL.0,
                    ptr::null_mut(),
                    audio_client.as_mut_ptr() as *mut _,
                )
                .ok()?;

            let audio_client = audio_client.assume_init();
            //let mut desired_format = mem::MaybeUninit::<*mut WAVEFORMATEXTENSIBLE>::zeroed();
            let desired_format = WAVEFORMATEX {
                cbSize: 0,
                nAvgBytesPerSec: 176400, 
                nBlockAlign: 4, 
                nChannels: 2,
                nSamplesPerSec: 44100,
                wBitsPerSample: 16,
                wFormatTag: WAVE_FORMAT_PCM as u16,
            };
            //let sample = WAVEFORMATEXTENSIBLE_0 {
            //    wValidBitsPerSample: 24,
            //};
            let desired_format_ex = WAVEFORMATEXTENSIBLE {
                Format: desired_format,
                Samples: WAVEFORMATEXTENSIBLE_0 {
                    wValidBitsPerSample: 16,
                },
                SubFormat: KSDATAFORMAT_SUBTYPE_PCM,
                dwChannelMask: 0,
            };

            let supported = audio_client.IsFormatSupported(AUDCLNT_SHAREMODE_EXCLUSIVE, &desired_format_ex as *const _ as *const WAVEFORMATEX, ptr::null_mut());
            println!("supported {:?}", supported.ok());
        }
    }
    println!("done");
    Ok(())
}
