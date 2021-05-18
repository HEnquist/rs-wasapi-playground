use std::mem;
use std::ptr;
use std::slice;
use wasapi::{
    PKEY_Device_FriendlyName,
    Windows::Win32::Media::Audio::CoreAudio::{
        eConsole, eRender, IAudioClient, IAudioRenderClient, IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator,
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
            let bits = 16;
            let validbits = 16;
            let srate = 48000;
            let channels = 2;
            let blockalign = channels * bits / 8;
            let byterate = srate * blockalign;

            let desired_format = WAVEFORMATEX {
                cbSize: 22,
                nAvgBytesPerSec: byterate,
                nBlockAlign: blockalign as u16, 
                nChannels: channels as u16,
                nSamplesPerSec: srate,
                wBitsPerSample: bits as u16,
                wFormatTag: WAVE_FORMAT_EXTENSIBLE as u16,
            };
            let sample = WAVEFORMATEXTENSIBLE_0 {
                wValidBitsPerSample: validbits,
            };
            let desired_format_ex = WAVEFORMATEXTENSIBLE {
                Format: desired_format,
                Samples: sample,
                SubFormat: KSDATAFORMAT_SUBTYPE_PCM,
                dwChannelMask: 3,
            };
            println!("nAvgBytesPerSec {:?}", { desired_format_ex.Format.nAvgBytesPerSec });
            println!("cbSize {:?}", { desired_format_ex.Format.cbSize });
            println!("nBlockAlign {:?}", { desired_format_ex.Format.nBlockAlign });
            println!("wBitsPerSample {:?}", { desired_format_ex.Format.wBitsPerSample });
            println!("nSamplesPerSec {:?}", { desired_format_ex.Format.nSamplesPerSec });
            println!("wFormatTag {:?}", { desired_format_ex.Format.wFormatTag });
            println!("wValidBitsPerSample {:?}", { desired_format_ex.Samples.wValidBitsPerSample });
            println!("SubFormat {:?}", { desired_format_ex.SubFormat });


            let desired_format_simple = WAVEFORMATEX {
                cbSize: 0,
                nAvgBytesPerSec: byterate,
                nBlockAlign: blockalign as u16, 
                nChannels: channels as u16,
                nSamplesPerSec: srate,
                wBitsPerSample: bits as u16,
                wFormatTag: WAVE_FORMAT_PCM as u16,
            };

            let supported = audio_client.IsFormatSupported(AUDCLNT_SHAREMODE_EXCLUSIVE, &desired_format_ex as *const _ as *const WAVEFORMATEX, ptr::null_mut());
            println!("supported {:?}\n", supported.ok());

            let mut def_time = 0;
            let mut min_time = 0;
            let res = audio_client.GetDevicePeriod(&mut def_time, &mut min_time);
            println!("result {:?}", res.ok());
            println!("default period {}, min period {}", def_time, min_time);


            let res = audio_client.Initialize(AUDCLNT_SHAREMODE_EXCLUSIVE,
                    AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                    def_time,
                    def_time,
                    &desired_format_ex as *const _ as *const WAVEFORMATEX,
                    std::ptr::null());
            println!("result {:?}", res.ok());

            let h_event = CreateEventA(std::ptr::null_mut(), false, false, PSTR::default());

            let hr = audio_client.SetEventHandle(h_event);
            println!("result {:?}", hr.ok());
            
            let mut bufferFrameCount = 0;
            let hr = audio_client.GetBufferSize(&mut bufferFrameCount);
            println!("result {:?}", hr.ok());
            println!("bufferFrameCount {}",bufferFrameCount);

            let render_client: IAudioRenderClient = audio_client.GetService()?;

            let hr = audio_client.Start();
            println!("result {:?}", hr.ok());

            for n in 0..128 {
                let mut data = mem::MaybeUninit::uninit();
                render_client
                    .GetBuffer(bufferFrameCount, data.as_mut_ptr())
                    .ok()?;

                let mut dataptr = data.assume_init();
                let mut databuf = slice::from_raw_parts_mut(dataptr, (bufferFrameCount*blockalign) as usize);
                for m in 0..bufferFrameCount*blockalign/2 {
                    databuf[m as usize] = 10;
                }
                render_client.ReleaseBuffer(bufferFrameCount, 0);
                println!("wrote frames");

                let retval = WaitForSingleObject(h_event, 100);
                if (retval != WAIT_OBJECT_0)
                {
                    // Event handle timed out after a 2-second wait.
                    audio_client.Stop();
                    break;
                }

            }

            /*
            let mut desired_format_result: *mut WAVEFORMATEXTENSIBLE = ptr::null_mut();
            let supported = audio_client.IsFormatSupported(AUDCLNT_SHAREMODE_EXCLUSIVE, &desired_format_ex as *const _ as *const WAVEFORMATEX, &mut desired_format_result as *mut _ as *mut *mut WAVEFORMATEX);
            println!("nAvgBytesPerSec {:?}", (*desired_format_result).Format.nAvgBytesPerSec);
            println!("cbSize {:?}", (*desired_format_result).Format.cbSize);
            println!("nBlockAlign {:?}", (*desired_format_result).Format.nBlockAlign);
            println!("wBitsPerSample {:?}", (*desired_format_result).Format.wBitsPerSample);
            println!("nSamplesPerSec {:?}", (*desired_format_result).Format.nSamplesPerSec);
            println!("wFormatTag {:?}", (*desired_format_result).Format.wFormatTag);
            println!("wValidBitsPerSample {:?}", (*desired_format_result).Samples.wValidBitsPerSample);
            println!("SubFormat {:?}", (*desired_format_result).SubFormat);
            */
        }

    }
    println!("done");
    Ok(())
}
