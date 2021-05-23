use std::collections::VecDeque;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::mpsc;
use std::sync::{Arc, Barrier, RwLock};
use std::thread;
use std::time;

use windows::initialize_mta;
use std::error;
use wasapi::wasapi::*;


type Res<T> = Result<T, Box<dyn error::Error>>;


fn main() -> Res<()> {
    initialize_mta()?;
    /*
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
    */
    let blockalign = 4;
    let (tx_dev, rx_dev): (std::sync::mpsc::SyncSender<Vec<u8>>, std::sync::mpsc::Receiver<Vec<u8>>) = mpsc::sync_channel(2);
    let buffer_fill = Arc::new(AtomicUsize::new(0));
    let buffer_fill_clone = buffer_fill.clone();
    
    // Playback
    let _handle = thread::Builder::new()
        .name("Player".to_string())
        .spawn(move || {
            let collection = DeviceCollection::new(false).unwrap();
            let device = collection.get_device_with_name("SPDIF Interface (FX-AUDIO-DAC-X6)").unwrap();
            let audio_client = device.get_iaudioclient().unwrap();

            let desired_format_ex = WaveFormat::new(16, 16, false, 48000, 2);
            let blockalign = desired_format_ex.get_blockalign();
            desired_format_ex.print_waveformat();

            let supported = audio_client.is_supported_exclusive(&desired_format_ex);
            println!("supported {:?}\n", supported);

            let (def_time, min_time) = audio_client.get_periods().unwrap();
            println!("default period {}, min period {}", def_time, min_time);


            audio_client.initialize_client(&desired_format_ex, def_time as i64).unwrap();

            let h_event = audio_client.set_get_eventhandle().unwrap();

            let buffer_frame_count = audio_client.get_bufferframecount().unwrap();

            let render_client = audio_client.get_audiorenderclient().unwrap();
            let mut sample_queue: VecDeque<u8> = VecDeque::with_capacity(100*blockalign as usize * (1024 + 2*buffer_frame_count as usize));
            audio_client.start_stream().unwrap();
            loop {
                //println!("deque len {}", sample_queue.len());
                while sample_queue.len() < (blockalign as usize * buffer_frame_count as usize) {
                    println!("need more samples");
                    
                    match rx_dev.recv_timeout(time::Duration::from_micros(1000)) {
                        Ok(chunk) => {
                            println!("got chunk");
                            for element in chunk.iter() {
                                sample_queue.push_back(*element);
                            }
                        }
                        Err(_) => {
                            println!("oops");
                            break;
                        }
                    }
                    println!("deque len2 {}", sample_queue.len());
                }
                //println!("wait for buf");

                //println!("write");
                render_client.write_to_device_from_deque(buffer_frame_count as usize, blockalign as usize, &mut sample_queue ).unwrap();
                if h_event.wait_for_event(10000).is_err() {
                    audio_client.stop_stream().unwrap();
                    break;
                }
            }
        });

    let mut timeval: usize = 0;
    for _n in 0..100 {
        let mut databuf = vec![0u8; (1024*blockalign) as usize];
        for m in 0..databuf.len() {
            databuf[m] = ((timeval%256)>128) as u8 * 10;
            timeval += 1;
        }
        println!("sending");
        tx_dev.send(databuf).unwrap();
        
    }
    //let device = get_device_with_name(&devs, "SPDIF Interface (FX-AUDIO-DAC-X6)")?;
    println!("done");
    Ok(())
}
