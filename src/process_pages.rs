use lopdf::{Document, Object, ObjectId, Stream, Dictionary};
use lopdf::content::Content;
use std::error::Error;

// ========== small helpers ==========
#[inline]
fn dict_get<'a>(dict: &'a Dictionary, key: &[u8]) -> Option<&'a Object> {
    dict.get(key).ok()
}

#[inline]
fn obj_as_dict_owned(obj: &Object, doc: &Document) -> Option<Dictionary> {
    match obj {
        Object::Dictionary(d) => Some(d.clone()),
        Object::Reference(idref) => {
            let d = doc.get_object(*idref).ok()?.as_dict().ok()?;
            Some(d.clone())
        }
        _ => None,
    }
}

// ========== public entry ==========
pub fn remove_blank_pages(doc: &mut Document) -> Result<(), Box<dyn Error>> {

    let page_ids: Vec<ObjectId> = doc.get_pages().values().cloned().collect();
    for pid in page_ids.into_iter().rev() {
        if page_is_blank(doc, pid)? {
            delete_page(doc, pid)?; // 정확 삭제
        }
    }

    doc.renumber_objects();

    // (있으면) 고아 객체 제거 -> 삭제된 페이지에서만 쓰이던 폰트/이미지도 제거됨
    let _ = doc.prune_objects();   // lopdf 0.38에 있으면 사용, 없으면 생략

    Ok(())
}

// ========== blank detection ==========
fn page_is_blank(doc: &mut Document, page_id: ObjectId) -> lopdf::Result<bool> {
    let streams = page_content_streams(doc, page_id)?;
    if streams.is_empty() {
        return Ok(true);
    }

    let resources = effective_resources(doc, page_id);

    for s in streams {
        let content = Content::decode(&s.content)?;
        if draws_something(doc, &content, &resources)? {
            return Ok(false);
        }
    }
    Ok(true)
}

fn draws_something(doc: &Document, content: &Content, resources: &Option<Dictionary>) -> lopdf::Result<bool> {
    for op in &content.operations {
        let name = op.operator.as_str();

        // 텍스트/경로/셰이딩/인라인 이미지
        if matches!(name, "Tj" | "TJ" | "'" | "\"" |
                          "S" | "s" | "f" | "F" | "f*" | "B" | "B*" | "b" | "b*" |
                          "sh" | "BI")
        {
            return Ok(true);
        }

        // XObject 호출 처리
        if name == "Do" {
            if let Some(first) = op.operands.get(0) {
                if let Some(res) = resources {
                    if let Some(xobjs_obj) = dict_get(res, b"XObject") {
                        let xdict = obj_as_dict_owned(xobjs_obj, doc).unwrap_or_else(Dictionary::new);

                        if let Object::Name(nm) = first {
                            if let Some(obj) = dict_get(&xdict, nm.as_slice()) {
                                if let Object::Reference(oid) = obj {
                                    let xobj = doc.get_object(*oid)?.as_stream()?;
                                    if let Some(sub_obj) = xobj.dict.get(b"Subtype").ok() {
                                        if let Object::Name(sub) = sub_obj {
                                            match sub.as_slice() {
                                                b"Image" => return Ok(true),
                                                b"Form"  => {
                                                    let inner = Content::decode(&xobj.content)?;
                                                    // Form 전용 Resources 우선
                                                    let frm_res = if let Some(r) = xobj.dict.get(b"Resources").ok() {
                                                        obj_as_dict_owned(r, doc)
                                                    } else {
                                                        resources.clone()
                                                    };
                                                    if draws_something(doc, &inner, &frm_res)? {
                                                        return Ok(true);
                                                    }
                                                }
                                                _ => {}
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    Ok(false)
}

fn page_content_streams(doc: &Document, page_id: ObjectId) -> lopdf::Result<Vec<Stream>> {
    let page = doc.get_object(page_id)?.as_dict()?;
    let mut out = Vec::new();

    if let Some(obj) = page.get(b"Contents").ok() {
        match obj {
            Object::Reference(cid) => {
                out.push(doc.get_object(*cid)?.as_stream()?.clone());
            }
            Object::Array(arr) => {
                for o in arr {
                    if let Object::Reference(id) = o {
                        out.push(doc.get_object(*id)?.as_stream()?.clone());
                    }
                }
            }
            Object::Stream(s) => out.push(s.clone()),
            _ => {}
        }
    }
    Ok(out)
}

fn effective_resources(doc: &Document, page_id: ObjectId) -> Option<Dictionary> {
    // 페이지에 직접 있으면 사용
    let page = doc.get_object(page_id).ok()?.as_dict().ok()?;
    if let Some(obj) = page.get(b"Resources").ok() {
        return obj_as_dict_owned(obj, doc);
    }

    // Parent 사슬 따라 상속 탐색
    let mut cur = page;
    loop {
        match dict_get(cur, b"Parent") {
            Some(Object::Reference(pid)) => {
                let parent = doc.get_object(*pid).ok()?.as_dict().ok()?;
                if let Some(obj) = parent.get(b"Resources").ok() {
                    return obj_as_dict_owned(obj, doc);
                }
                cur = parent;
            }
            _ => break,
        }
    }
    None
}

// ========== deletion ==========
fn delete_page(doc: &mut Document, page_id: ObjectId) -> Result<(), Box<dyn Error>> {
    // 1) Parent
    let parent_id = {
        let page_dict = doc.get_object(page_id)?.as_dict()?;
        match page_dict.get(b"Parent")? {
            Object::Reference(pid) => *pid,
            _ => return Err("page has no Parent".into()),
        }
    };

    // 2) Parent.Kids에서 제거
    {
        let parent = doc.get_object_mut(parent_id)?;
        let pdict = parent.as_dict_mut()?;
        if let Ok(obj) = pdict.get_mut(b"Kids") {
            if let Object::Array(kids) = obj {
                kids.retain(|o| !matches!(o, Object::Reference(id) if *id == page_id));
            }
        }
    }

    // 3) 조상 /Count 감소
    let mut cur = Some(parent_id);
    while let Some(pid) = cur {
        let pobj = doc.get_object_mut(pid)?;
        let pdict = pobj.as_dict_mut()?;

        if let Ok(obj) = pdict.get_mut(b"Count") {
            if let Object::Integer(c) = obj { // c: &mut i64 (match ergonomics)
                *c -= 1;
            }
        }

        cur = match pdict.get(b"Parent") {
            Ok(Object::Reference(pp)) => Some(*pp),
            _ => None,
        };
    }

    // 4) 삭제할 Contents 참조들 먼저 수집 (borrow 충돌 방지)
    let content_ids: Vec<ObjectId> = {
        let pd = doc.get_object(page_id)?.as_dict()?;
        let mut ids = Vec::new();
        if let Some(obj) = pd.get(b"Contents").ok() {
            match obj {
                Object::Reference(cid) => ids.push(*cid),
                Object::Array(arr) => {
                    for o in arr {
                        if let Object::Reference(cid) = o {
                            ids.push(*cid);
                        }
                    }
                }
                _ => {}
            }
        }
        ids
    };

    // 5) Contents 스트림 제거
    for cid in content_ids {
        doc.objects.remove(&cid);
    }

    // 6) 페이지 객체 제거
    doc.objects.remove(&page_id);

    Ok(())
}

// Helvetica / Helvetica-Bold (WinAnsi, 32..126) widths in 1/1000 em
pub fn stamp_watermarks(doc: &mut Document) -> Result<(), Box<dyn Error>> {
    // 1) 공유 리소스: Helvetica-Bold / 반투명 GState
    let font_id = {
        let mut d = Dictionary::new();
        d.set("Type", "Font");
        d.set("Subtype", "Type1");
        d.set("BaseFont", "Helvetica-Bold");
        let id = doc.new_object_id();
        doc.objects.insert(id, Object::Dictionary(d));
        id
    };
    let gs_id = {
        let mut d = Dictionary::new();
        d.set("Type", "ExtGState");
        d.set("BM", "Normal");
        d.set("ca", Object::Real(0.18));
        d.set("CA", Object::Real(0.18));
        let id = doc.new_object_id();
        doc.objects.insert(id, Object::Dictionary(d));
        id
    };

    // 헬베티카 폭표(중앙정렬용)
    const HELV_W_32_126: [i16; 95] = [
        278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,
        556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,
        1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,
        667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,
        333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,
        556,556,333,500,278,556,500,722,500,500,500,334,260,334,584,
    ];
    let text_width = |s: &str, fs: f64| -> f64 {
        let w1000: f64 = s.bytes().map(|b|
            if (32..=126).contains(&b) { HELV_W_32_126[(b-32) as usize] as f64 } else { 600.0 }
        ).sum();
        w1000 * fs / 1000.0
    };

    let page_ids: Vec<ObjectId> = doc.get_pages().values().cloned().collect();
    for pid in page_ids {
        // --- 페이지 박스/중앙 ---
        let (llx, lly, urx, ury) = effective_mediabox(doc, pid).ok_or("Page has no MediaBox")?;
        let (w, h) = (urx - llx, ury - lly);
        let (cx, cy) = (llx + w/2.0, lly + h/2.0);

        // --- 기존 컨텐츠 스트림 모으기(복사본) ---
        let old_streams = page_content_streams(doc, pid)?;
        // 내용이 비어있지 않으면 → Form XObject로 감싸기
        let mut xobj_name_for_old: Option<Vec<u8>> = None;
        if !old_streams.is_empty() {
            // 1) 기존 컨텐츠 바이트 결합
            let mut concat = Vec::<u8>::new();
            for s in &old_streams { concat.extend_from_slice(&s.content); concat.push(b'\n'); }

            // 2) Form XObject 생성(기존 리소스를 폼 안으로)
            let mut form_dict = Dictionary::new();
            form_dict.set("Type", "XObject");
            form_dict.set("Subtype", "Form");
            form_dict.set("FormType", 1);
            form_dict.set("BBox", Object::Array(vec![llx.into(), lly.into(), urx.into(), ury.into()]));
            if let Some(res) = effective_resources(doc, pid) {
                form_dict.set("Resources", Object::Dictionary(res));
            }
            let form_id = {
                let id = doc.new_object_id();
                doc.objects.insert(id, Object::Stream(Stream::new(form_dict, concat)));
                id
            };

            // 3) 페이지 리소스의 /XObject에 등록할 이름
            xobj_name_for_old = Some(b"OLD_FORM".to_vec());

            // 4) 페이지 리소스 사본 만들고 /XObject에 OLD_FORM 추가(+ 우리 폰트/GS)
            let mut resources = {
                let page_ro = doc.get_object(pid)?.as_dict()?.clone();
                if let Some(obj) = page_ro.get(b"Resources").ok() {
                    obj_as_dict_owned(obj, doc).unwrap_or_else(Dictionary::new)
                } else { Dictionary::new() }
            };
            // /XObject
            let mut xobjs = if let Some(o) = resources.get(b"XObject").ok() {
                obj_as_dict_owned(o, doc).unwrap_or_else(Dictionary::new)
            } else { Dictionary::new() };
            xobjs.set("OLD_FORM", Object::Reference(form_id));
            resources.set("XObject", Object::Dictionary(xobjs));
            // /Font
            let mut fr = if let Some(o) = resources.get(b"Font").ok() {
                obj_as_dict_owned(o, doc).unwrap_or_else(Dictionary::new)
            } else { Dictionary::new() };
            fr.set("F_ARC", Object::Reference(font_id));
            resources.set("Font", Object::Dictionary(fr));
            // /ExtGState
            let mut gs = if let Some(o) = resources.get(b"ExtGState").ok() {
                obj_as_dict_owned(o, doc).unwrap_or_else(Dictionary::new)
            } else { Dictionary::new() };
            gs.set("GS_ARC", Object::Reference(gs_id));
            resources.set("ExtGState", Object::Dictionary(gs));

            // 페이지에 리소스 적용(가변 대여 한 번)
            {
                let page_mut = doc.get_object_mut(pid)?;
                let pd = page_mut.as_dict_mut()?;
                pd.set("Resources", Object::Dictionary(resources));
            }
        } else {
            // 기존 리소스가 없어도 워터마크용 Font/GS는 필요
            let mut resources = {
                let page_ro = doc.get_object(pid)?.as_dict()?.clone();
                if let Some(obj) = page_ro.get(b"Resources").ok() {
                    obj_as_dict_owned(obj, doc).unwrap_or_else(Dictionary::new)
                } else { Dictionary::new() }
            };
            let mut fr = if let Some(o) = resources.get(b"Font").ok() {
                obj_as_dict_owned(o, doc).unwrap_or_else(Dictionary::new)
            } else { Dictionary::new() };
            fr.set("F_ARC", Object::Reference(font_id));
            resources.set("Font", Object::Dictionary(fr));
            let mut gs = if let Some(o) = resources.get(b"ExtGState").ok() {
                obj_as_dict_owned(o, doc).unwrap_or_else(Dictionary::new)
            } else { Dictionary::new() };
            gs.set("GS_ARC", Object::Reference(gs_id));
            resources.set("ExtGState", Object::Dictionary(gs));
            {
                let page_mut = doc.get_object_mut(pid)?;
                let pd = page_mut.as_dict_mut()?;
                pd.set("Resources", Object::Dictionary(resources));
            }
        }

        // --- 워터마크 텍스트(중앙정렬 + 진짜/가짜 볼드 + 밑줄) ---
        let text = "ARC";
        let fs = 0.25 * w.min(h);
        let tw = text_width(text, fs);
        let theta = 45f64.to_radians(); let (c,s) = (theta.cos(), theta.sin()); let ms = -s;
        let dx = -tw/2.0;          // 정확 중앙 정렬
        let dy = -(fs*0.35);
        let stroke_w = fs*0.060;
        let ul_th = fs*0.050;
        let ul_off = fs*0.180;
        let udy = dy - ul_off;

        let mut contents_refs: Vec<Object> = Vec::new();

        // 1) 기존 컨텐츠 폼 그리기(그래픽 상태 끌어안고 그 안에서만 영향)
        if let Some(name) = xobj_name_for_old.clone() {
            let draw_old = format!(
                "q\n/{name} Do\nQ\n",
                name = String::from_utf8_lossy(&name)
            );
            let draw_old_id = doc.new_object_id();
            doc.objects.insert(draw_old_id, Object::Stream(Stream::new(Dictionary::new(), draw_old.into_bytes())));
            contents_refs.push(Object::Reference(draw_old_id));
        }

        // 2) 그 위에 워터마크
        let wm_stream = format!(
            concat!(
                "q\n",
                "/GS_ARC gs\n",
                "1 0 0 rg  1 0 0 RG\n",
                "{c} {s} {ms} {c} {cx} {cy} cm\n",
                "BT\n/F_ARC {fs:.3} Tf\n{dx:.3} {dy:.3} Td\n2 Tr {sw:.3} w\n({text}) Tj\nET\n",
                "1 0 0 1 {dx:.3} {udy:.3} cm\n{ul:.3} w\n0 0 m {tw:.3} 0 l S\n",
                "Q\n"
            ),
            c=c, s=s, ms=ms, cx=cx, cy=cy, fs=fs, dx=dx, dy=dy,
            sw=stroke_w, udy=udy, ul=ul_th, tw=tw, text=text
        );
        let wm_id = doc.new_object_id();
        doc.objects.insert(wm_id, Object::Stream(Stream::new(Dictionary::new(), wm_stream.into_bytes())));
        contents_refs.push(Object::Reference(wm_id));

        // 3) 페이지의 Contents 교체
        {
            let page_mut = doc.get_object_mut(pid)?;
            let pd = page_mut.as_dict_mut()?;
            if contents_refs.len() == 1 {
                pd.set("Contents", contents_refs.remove(0));
            } else {
                pd.set("Contents", Object::Array(contents_refs));
            }
        }
    }

    Ok(())
}



fn as_f64(n: &Object) -> Option<f64> {
    match n {
        Object::Integer(i) => Some(*i as f64),
        Object::Real(r) => Some(*r as f64),
        _ => None,
    }
}

fn effective_mediabox(doc: &Document, page_id: ObjectId) -> Option<(f64, f64, f64, f64)> {
    // 페이지에서 시작해 Parent 체인을 올라가며 /MediaBox 탐색
    let mut cur = doc.get_object(page_id).ok()?.as_dict().ok()?;
    loop {
        if let Some(obj) = cur.get(b"MediaBox").ok() {
            if let Object::Array(a) = obj {
                if a.len() == 4 {
                    let llx = as_f64(&a[0])?;
                    let lly = as_f64(&a[1])?;
                    let urx = as_f64(&a[2])?;
                    let ury = as_f64(&a[3])?;
                    return Some((llx, lly, urx, ury));
                }
            }
        }
        match dict_get(cur, b"Parent") {
            Some(Object::Reference(pid)) => {
                cur = doc.get_object(*pid).ok()?.as_dict().ok()?;
            }
            _ => break,
        }
    }
    None
}


pub fn post_process_arc(doc: &mut Document) -> Result<(), Box<dyn Error>> {
    let _ = doc.decompress();
    remove_blank_pages(doc)?;
    stamp_watermarks(doc)?;
    _ = doc.compress();
    Ok(())
}