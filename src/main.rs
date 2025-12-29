use std::{
    mem::MaybeUninit,
    ops::Range,
    sync::{
        Arc, Mutex,
        atomic::{AtomicU32, Ordering},
    },
};

use windows::Win32::{
    Foundation::{E_FAIL, E_POINTER, POINT, RECT, S_OK},
    System::Com::{CLSCTX_INPROC_SERVER, CoCreateInstance, FORMATETC, IDataObject},
    UI::TextServices::{
        CLSID_TF_ThreadMgr, ITextStoreACP2, ITextStoreACP2_Impl, ITextStoreACPSink,
        ITfCandidateListUIElement, ITfCompositionView, ITfContextOwnerCompositionSink,
        ITfContextOwnerCompositionSink_Impl, ITfDocumentMgr, ITfRange, ITfSource, ITfThreadMgrEx,
        ITfUIElementMgr, ITfUIElementSink, ITfUIElementSink_Impl, TF_IPP_CAPS_UIELEMENTENABLED,
        TF_TF_IGNOREEND, TS_AE_END, TS_ATTRVAL, TS_E_NOLAYOUT, TS_LF_READ, TS_LF_READWRITE,
        TS_RT_PLAIN, TS_RUNINFO, TS_SELECTION_ACP, TS_SELECTIONSTYLE, TS_STATUS, TS_TEXTCHANGE,
    },
};
use windows::core::{
    BOOL, GUID, HRESULT, IUnknown, Interface, PCWSTR, PWSTR, Ref, Result, implement,
};
use winit::{
    application::ApplicationHandler,
    event::WindowEvent,
    event_loop::{ActiveEventLoop, ControlFlow, EventLoop},
    window::{Window, WindowId},
};

#[derive(Default)]
struct App {
    window: Option<Window>,
    thread_mgr: Option<ITfThreadMgrEx>,
    document_mgr: Option<ITfDocumentMgr>,
}

impl ApplicationHandler for App {
    fn resumed(&mut self, event_loop: &ActiveEventLoop) {
        let window = event_loop
            .create_window(Window::default_attributes())
            .unwrap();
        // window.set_ime_allowed(true);
        unsafe {
            let thread_mgr: ITfThreadMgrEx =
                CoCreateInstance(&CLSID_TF_ThreadMgr, None, CLSCTX_INPROC_SERVER).unwrap();
            // let thread_mgr_ex: ITfThreadMgrEx = thread_mgr.cast().unwrap();

            let mut client_id = MaybeUninit::uninit();
            thread_mgr
                .ActivateEx(client_id.as_mut_ptr(), TF_IPP_CAPS_UIELEMENTENABLED)
                .unwrap();
            let client_id = client_id.assume_init();

            let ui_element_mgr: ITfUIElementMgr = thread_mgr.cast().unwrap();
            let ui_element_source: ITfSource = ui_element_mgr.cast().unwrap();

            let ui_element_sink: ITfUIElementSink = UIElementSink { ui_element_mgr }.into();
            ui_element_source
                .AdviseSink(&ITfUIElementSink::IID, &ui_element_sink)
                .unwrap();

            let document_mgr: ITfDocumentMgr = thread_mgr.CreateDocumentMgr().unwrap();

            let mut context = MaybeUninit::uninit();
            let mut edit_cookie = MaybeUninit::uninit();

            let edit_cookie_arc = Arc::new(AtomicU32::new(0));
            let composition_sink: ITfContextOwnerCompositionSink = TextStore {
                edit_cookie: edit_cookie_arc.clone(),
                sink: Mutex::new(Vec::new()),
                content: Mutex::new(Content {
                    text: Vec::new(),
                    selection: 0..0,
                }),
            }
            .into();
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
            println!("Edit cookie: {edit_cookie}");
            edit_cookie_arc.store(edit_cookie, Ordering::Release);

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

#[implement(ITfContextOwnerCompositionSink, ITextStoreACP2)]
struct TextStore {
    edit_cookie: Arc<AtomicU32>,
    sink: Mutex<Vec<ITextStoreACPSink>>,
    content: Mutex<Content>,
}

struct Content {
    text: Vec<u16>,
    selection: Range<i32>,
}

impl ITfContextOwnerCompositionSink_Impl for TextStore_Impl {
    fn OnStartComposition(&self, _pcomposition: Ref<ITfCompositionView>) -> Result<BOOL> {
        println!("OnStartComposition");
        Ok(true.into())
    }

    fn OnUpdateComposition(
        &self,
        _pcomposition: Ref<ITfCompositionView>,
        _prangenew: Ref<ITfRange>,
    ) -> Result<()> {
        println!("OnUpdateComposition");
        Ok(())
    }

    fn OnEndComposition(&self, pcomposition: Ref<ITfCompositionView>) -> Result<()> {
        println!("OnEndComposition");
        if let Some(composition) = pcomposition.as_ref() {
            unsafe {
                let range = composition.GetRange()?;
                let mut buffer = [0; 128];
                let mut chars = 0;
                let edit_cookie = self.edit_cookie.load(Ordering::Acquire);
                println!("Edit cookie: {edit_cookie}");
                range.GetText(edit_cookie, TF_TF_IGNOREEND, &mut buffer, &raw mut chars)?;
                println!(
                    "String: {:?}",
                    String::from_utf16(&buffer[..chars as usize])
                );
            }
        }
        Ok(())
    }
}

impl ITextStoreACP2_Impl for TextStore_Impl {
    fn AdviseSink(&self, riid: *const GUID, punk: Ref<IUnknown>, dwmask: u32) -> Result<()> {
        let mut sinks = self.sink.lock().unwrap();
        sinks.push(punk.clone().ok_or(E_POINTER)?.cast()?);
        Ok(())
    }

    fn UnadviseSink(&self, punk: Ref<IUnknown>) -> Result<()> {
        let mut sinks = self.sink.lock().unwrap();
        let target = punk.as_ref().ok_or(E_POINTER)?;
        sinks.retain(|x| x.cast::<IUnknown>().unwrap().eq(target));
        Ok(())
    }

    fn RequestLock(&self, dwlockflags: u32) -> Result<HRESULT> {
        let sinks = self.sink.lock().unwrap().clone();
        if dwlockflags & TS_LF_READWRITE.0 != 0 {
            for sink in &sinks {
                unsafe {
                    sink.OnLockGranted(TS_LF_READWRITE)?;
                }
            }
        } else if dwlockflags & TS_LF_READ.0 != 0 {
            for sink in &sinks {
                unsafe {
                    sink.OnLockGranted(TS_LF_READ)?;
                }
            }
        }
        Ok(S_OK)
    }

    fn GetStatus(&self) -> Result<TS_STATUS> {
        Ok(TS_STATUS {
            dwDynamicFlags: 0,
            dwStaticFlags: 0,
        })
    }

    fn QueryInsert(
        &self,
        acpteststart: i32,
        acptestend: i32,
        cch: u32,
        pacpresultstart: *mut i32,
        pacpresultend: *mut i32,
    ) -> Result<()> {
        todo!()
    }

    fn GetSelection(
        &self,
        ulindex: u32,
        ulcount: u32,
        pselection: *mut TS_SELECTION_ACP,
        pcfetched: *mut u32,
    ) -> Result<()> {
        let content = self.content.lock().unwrap();
        unsafe {
            *pselection = TS_SELECTION_ACP {
                acpStart: content.selection.start,
                acpEnd: content.selection.end,
                style: TS_SELECTIONSTYLE {
                    ase: TS_AE_END,
                    fInterimChar: false.into(),
                },
            };
            *pcfetched = 1;
        }
        Ok(())
    }

    fn SetSelection(&self, ulcount: u32, pselection: *const TS_SELECTION_ACP) -> Result<()> {
        if ulcount != 1 {
            return Err(E_FAIL.into());
        }
        let mut content = self.content.lock().unwrap();
        unsafe {
            content.selection.start = (*pselection).acpStart;
            content.selection.end = (*pselection).acpEnd;
        }
        Ok(())
    }

    fn GetText(
        &self,
        acpstart: i32,
        acpend: i32,
        pchplain: PWSTR,
        cchplainreq: u32,
        pcchplainret: *mut u32,
        prgruninfo: *mut TS_RUNINFO,
        cruninforeq: u32,
        pcruninforet: *mut u32,
        pacpnext: *mut i32,
    ) -> Result<()> {
        let content = self.content.lock().unwrap();
        println!("start: {acpstart}, end: {acpend}");
        let mut slice = &content.text[acpstart as usize..];
        if acpend != -1 {
            slice = &slice[..acpend as usize];
        }
        let count = slice.len().min(cchplainreq as usize);
        unsafe {
            pchplain.0.copy_from(slice.as_ptr(), count);
            *pcchplainret = count as u32;
            if cruninforeq > 0 {
                *prgruninfo = TS_RUNINFO {
                    uCount: *pcchplainret,
                    r#type: TS_RT_PLAIN,
                };
                *pcruninforet = 1;
            }
            *pacpnext = acpstart + (count as i32);
        }
        Ok(())
    }

    fn SetText(
        &self,
        dwflags: u32,
        acpstart: i32,
        acpend: i32,
        pchtext: &PCWSTR,
        cch: u32,
    ) -> Result<TS_TEXTCHANGE> {
        let mut content = self.content.lock().unwrap();
        let insert = (0..cch).map(|i| unsafe { *pchtext.0.add(i as usize) });
        content
            .text
            .splice(acpstart as usize..acpend as usize, insert);
        Ok(TS_TEXTCHANGE {
            acpStart: acpstart,
            acpOldEnd: acpend,
            acpNewEnd: acpstart + cch as i32,
        })
    }

    fn GetFormattedText(&self, acpstart: i32, acpend: i32) -> Result<IDataObject> {
        let mut content = self.content.lock().unwrap();
        todo!()
    }

    fn GetEmbedded(
        &self,
        acppos: i32,
        rguidservice: *const GUID,
        riid: *const GUID,
    ) -> Result<IUnknown> {
        todo!()
    }

    fn QueryInsertEmbedded(
        &self,
        pguidservice: *const GUID,
        pformatetc: *const FORMATETC,
    ) -> Result<BOOL> {
        todo!()
    }

    fn InsertEmbedded(
        &self,
        dwflags: u32,
        acpstart: i32,
        acpend: i32,
        pdataobject: Ref<IDataObject>,
    ) -> Result<TS_TEXTCHANGE> {
        todo!()
    }

    fn InsertTextAtSelection(
        &self,
        dwflags: u32,
        pchtext: &PCWSTR,
        cch: u32,
        pacpstart: *mut i32,
        pacpend: *mut i32,
        pchange: *mut TS_TEXTCHANGE,
    ) -> Result<()> {
        todo!()
    }

    fn InsertEmbeddedAtSelection(
        &self,
        dwflags: u32,
        pdataobject: Ref<IDataObject>,
        pacpstart: *mut i32,
        pacpend: *mut i32,
        pchange: *mut TS_TEXTCHANGE,
    ) -> Result<()> {
        todo!()
    }

    fn RequestSupportedAttrs(
        &self,
        dwflags: u32,
        cfilterattrs: u32,
        pafilterattrs: *const GUID,
    ) -> Result<()> {
        println!(
            "flags: {dwflags}, filter_attrs: {cfilterattrs}, filter_attrs: {:?}",
            unsafe { *pafilterattrs }
        );
        Ok(())
    }

    fn RequestAttrsAtPosition(
        &self,
        acppos: i32,
        cfilterattrs: u32,
        pafilterattrs: *const GUID,
        dwflags: u32,
    ) -> Result<()> {
        todo!()
    }

    fn RequestAttrsTransitioningAtPosition(
        &self,
        acppos: i32,
        cfilterattrs: u32,
        pafilterattrs: *const GUID,
        dwflags: u32,
    ) -> Result<()> {
        todo!()
    }

    fn FindNextAttrTransition(
        &self,
        acpstart: i32,
        acphalt: i32,
        cfilterattrs: u32,
        pafilterattrs: *const GUID,
        dwflags: u32,
        pacpnext: *mut i32,
        pffound: *mut BOOL,
        plfoundoffset: *mut i32,
    ) -> Result<()> {
        todo!()
    }

    fn RetrieveRequestedAttrs(
        &self,
        ulcount: u32,
        paattrvals: *mut TS_ATTRVAL,
        pcfetched: *mut u32,
    ) -> Result<()> {
        unsafe {
            *pcfetched = 0;
        }
        Ok(())
    }

    fn GetEndACP(&self) -> Result<i32> {
        todo!()
    }

    fn GetActiveView(&self) -> Result<u32> {
        Ok(0)
    }

    fn GetACPFromPoint(&self, vcview: u32, ptscreen: *const POINT, dwflags: u32) -> Result<i32> {
        todo!()
    }

    fn GetTextExt(
        &self,
        vcview: u32,
        acpstart: i32,
        acpend: i32,
        prc: *mut RECT,
        pfclipped: *mut BOOL,
    ) -> Result<()> {
        Err(TS_E_NOLAYOUT.into())
    }

    fn GetScreenExt(&self, vcview: u32) -> Result<RECT> {
        println!("view: {vcview}");
        Ok(RECT {
            left: 0,
            top: 0,
            right: 100,
            bottom: 50,
        })
    }
}

#[implement(ITfUIElementSink)]
struct UIElementSink {
    ui_element_mgr: ITfUIElementMgr,
}

impl ITfUIElementSink_Impl for UIElementSink_Impl {
    fn BeginUIElement(&self, dwuielementid: u32, pbshow: *mut BOOL) -> Result<()> {
        println!("BeginUIElement: {dwuielementid}");
        unsafe {
            let ui_element = self.ui_element_mgr.GetUIElement(dwuielementid)?;
            println!("UIElement: {:?}", ui_element.GetGUID());
            *pbshow = false.into()
        }
        Ok(())
    }

    fn UpdateUIElement(&self, dwuielementid: u32) -> Result<()> {
        println!("UpdateUIElement: {dwuielementid}");
        unsafe {
            let ui_element = self.ui_element_mgr.GetUIElement(dwuielementid)?;
            if let Ok(candidate_list) = ui_element.cast::<ITfCandidateListUIElement>() {
                let count = candidate_list.GetCount()?;
                let mut list = Vec::with_capacity(count as usize);
                for i in 0..count {
                    list.push(candidate_list.GetString(i)?);
                }
                println!("{list:?}");
            }
        }
        Ok(())
    }

    fn EndUIElement(&self, dwuielementid: u32) -> Result<()> {
        println!("UpdateUIElement: {dwuielementid}");
        Ok(())
    }
}
