use std::mem::MaybeUninit;

use windows::Win32::System::Com::{CLSCTX_INPROC_SERVER, CoCreateInstance};
use windows::Win32::UI::TextServices::{
    CLSID_TF_ThreadMgr, ITfCandidateListUIElement, ITfCompositionView,
    ITfContextOwnerCompositionSink, ITfContextOwnerCompositionSink_Impl, ITfDocumentMgr, ITfSource,
    ITfThreadMgr, ITfThreadMgrEx, ITfUIElementMgr, ITfUIElementSink, ITfUIElementSink_Impl,
    TF_IPP_CAPS_UIELEMENTENABLED,
};
use windows::core::{BOOL, Interface, Ref, Result, implement};
use winit::application::ApplicationHandler;
use winit::event::WindowEvent;
use winit::event_loop::{ActiveEventLoop, ControlFlow, EventLoop};
use winit::platform::windows::WindowExtWindows;
use winit::raw_window_handle::RawWindowHandle;
use winit::window::{Window, WindowId};

#[derive(Default)]
struct App {
    window: Option<Window>,
    thread_mgr: Option<ITfThreadMgr>,
    document_mgr: Option<ITfDocumentMgr>,
}

impl ApplicationHandler for App {
    fn resumed(&mut self, event_loop: &ActiveEventLoop) {
        let window = event_loop
            .create_window(Window::default_attributes())
            .unwrap();
        // window.set_ime_allowed(true);
        unsafe {
            let thread_mgr: ITfThreadMgr =
                CoCreateInstance(&CLSID_TF_ThreadMgr, None, CLSCTX_INPROC_SERVER).unwrap();
            let thread_mgr_ex: ITfThreadMgrEx = thread_mgr.cast().unwrap();

            let mut client_id = MaybeUninit::uninit();
            thread_mgr_ex
                .ActivateEx(client_id.as_mut_ptr(), TF_IPP_CAPS_UIELEMENTENABLED)
                .unwrap();
            let client_id = client_id.assume_init();

            let ui_element_mgr: ITfUIElementMgr = thread_mgr_ex.cast().unwrap();
            let ui_element_source: ITfSource = ui_element_mgr.cast().unwrap();

            let ui_element_sink: ITfUIElementSink = UIElementSink { ui_element_mgr }.into();
            ui_element_source
                .AdviseSink(&ITfUIElementSink::IID, &ui_element_sink)
                .unwrap();

            let document_mgr: ITfDocumentMgr = thread_mgr_ex.CreateDocumentMgr().unwrap();

            let mut context = MaybeUninit::uninit();
            let mut edit_cookie = MaybeUninit::uninit();

            let composition_sink: ITfContextOwnerCompositionSink = CompositionSink.into();
            document_mgr
                .CreateContext(
                    client_id,
                    0,
                    &composition_sink,
                    context.as_mut_ptr(),
                    edit_cookie.as_mut_ptr(),
                )
                .unwrap();
            let context = context.assume_init().unwrap();
            let edit_cookie = edit_cookie.assume_init();
            println!("edit cookie: {edit_cookie}");

            document_mgr.Push(&context).unwrap();

            self.thread_mgr = Some(thread_mgr);
            self.document_mgr = Some(document_mgr);
        }
        self.window = Some(window);
    }

    fn window_event(&mut self, event_loop: &ActiveEventLoop, id: WindowId, event: WindowEvent) {
        match event {
            WindowEvent::CloseRequested => {
                println!("The close button was pressed; stopping");
                event_loop.exit();
            }
            WindowEvent::RedrawRequested => {
                // Redraw the application.
                //
                // It's preferable for applications that do not render continuously to render in
                // this event rather than in AboutToWait, since rendering in here allows
                // the program to gracefully handle redraws requested by the OS.

                // Draw.

                // Queue a RedrawRequested event.
                //
                // You only need to call this if you've determined that you need to redraw in
                // applications which do not always need to. Applications that redraw continuously
                // can render here instead.
                // self.window.as_ref().unwrap().request_redraw();
            }
            WindowEvent::Ime(ime_event) => {
                println!("Ime Event: {ime_event:?}");
            }
            WindowEvent::Focused(focused) => unsafe {
                let hwnd: u64 = id.into();
                println!("HWND: {hwnd}");
                let _ = self
                    .thread_mgr
                    .as_ref()
                    .unwrap()
                    .AssociateFocus(
                        std::mem::transmute(hwnd),
                        if focused {
                            self.document_mgr.as_ref()
                        } else {
                            None
                        },
                    )
                    .inspect(|x| println!("AssociateFocus: {x:?}"))
                    .inspect_err(|x| println!("AssociateFocus: {x}"));
            },
            _ => (),
        }
    }
}

fn main() {
    let event_loop = EventLoop::new().unwrap();

    // ControlFlow::Poll continuously runs the event loop, even if the OS hasn't
    // dispatched any events. This is ideal for games and similar applications.
    // event_loop.set_control_flow(ControlFlow::Poll);

    // ControlFlow::Wait pauses the event loop if no events are available to process.
    // This is ideal for non-game applications that only update in response to user
    // input, and uses significantly less power/CPU time than ControlFlow::Poll.
    event_loop.set_control_flow(ControlFlow::Wait);

    let mut app = App::default();
    event_loop.run_app(&mut app).unwrap();
}

#[implement(ITfContextOwnerCompositionSink)]
struct CompositionSink;

impl ITfContextOwnerCompositionSink_Impl for CompositionSink_Impl {
    fn OnStartComposition(&self, _pcomposition: Ref<ITfCompositionView>) -> Result<BOOL> {
        println!("OnStartComposition");
        Ok(true.into())
    }

    fn OnUpdateComposition(
        &self,
        _pcomposition: Ref<ITfCompositionView>,
        _prangenew: Ref<windows::Win32::UI::TextServices::ITfRange>,
    ) -> Result<()> {
        println!("OnUpdateComposition");
        Ok(())
    }

    fn OnEndComposition(&self, _pcomposition: Ref<ITfCompositionView>) -> Result<()> {
        println!("OnEndComposition");
        Ok(())
    }
}

#[implement(ITfUIElementSink)]
struct UIElementSink {
    ui_element_mgr: ITfUIElementMgr,
}

impl ITfUIElementSink_Impl for UIElementSink_Impl {
    fn BeginUIElement(
        &self,
        dwuielementid: u32,
        pbshow: *mut windows_core::BOOL,
    ) -> windows_core::Result<()> {
        println!("BeginUIElement: {dwuielementid}");
        unsafe {
            let ui_element = self.ui_element_mgr.GetUIElement(dwuielementid).unwrap();
            println!("UIElement: {:?}", ui_element.GetGUID());
            *pbshow = false.into()
        }
        Ok(())
    }

    fn UpdateUIElement(&self, dwuielementid: u32) -> windows_core::Result<()> {
        println!("UpdateUIElement: {dwuielementid}");
        unsafe {
            let ui_element = self.ui_element_mgr.GetUIElement(dwuielementid).unwrap();
            if let Ok(candidate_list) = ui_element.cast::<ITfCandidateListUIElement>() {
                let count = candidate_list.GetCount().unwrap();
                let mut list = Vec::with_capacity(count as usize);
                for i in 0..count {
                    list.push(candidate_list.GetString(i).unwrap());
                }
                println!("{list:?}");
            }
        }
        Ok(())
    }

    fn EndUIElement(&self, dwuielementid: u32) -> windows_core::Result<()> {
        println!("UpdateUIElement: {dwuielementid}");
        Ok(())
    }
}
