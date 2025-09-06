use clap::Parser;
use glob::glob;
use lopdf::{Document, Object, ObjectId, Stream};
use lopdf::dictionary;
use pdf_writer::{Content, Finish, Name, Pdf, Rect, Ref as PdfRef};
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};
mod process_pages;

/// Bind front + SVGs + back into a single PDF (vector)
#[derive(Parser, Debug)]
#[command(name="pdf_bind", about="Bind front + SVGs + back into a single PDF")]
struct Args {
    /// Target page width (default: 8.5)
    #[arg(long, default_value_t = 8.5)]
    width: f64,
    /// Target page height (default: 8.5)
    #[arg(long, default_value_t = 8.5)]
    height: f64,
    /// Unit type: "in" or "cm" (default: in)
    #[arg(long, default_value = "in")]
    r#type: String,
    /// If true, and front_matter page count is odd, insert a blank page to make it even
    #[arg(long, default_value_t = false)]
    make_even: bool,
    /// ARC mode: true => DO NOT insert blanks between SVG pages; false => insert 1 blank between SVG pages
    #[arg(long, default_value_t = false)]
    arc: bool,
}

impl Args {
    /// Book preset: 8.5x8.5 inches, make_even=false, arc=false
    pub fn book() -> Args {
        Args {
            width: 8.5,
            height: 8.5,
            r#type: "in".to_string(),
            make_even: false,
            arc: false,
        }
    }

    /// ARC preset: 8.5x8.5 inches, make_even=false, arc=true
    pub fn arc() -> Args {
        Args {
            width: 8.5,
            height: 8.5,
            r#type: "in".to_string(),
            make_even: false,
            arc: true,
        }
    }
}

fn to_points(v: f64, unit: &str) -> f64 {
    match unit.to_ascii_lowercase().as_str() {
        "cm" => v / 2.54 * 72.0,
        "in" | "inch" | "inches" => v * 72.0,
        _ => v * 72.0,
    }
}

/// xref 안정화를 위해 입력 PDF를 로드 후 곧바로 저장
fn roundtrip_save(input: &Path, out: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut doc = Document::load(input)?;
    doc.save(out)?;
    Ok(())
}

fn pages_root_id(doc: &Document) -> Result<ObjectId, Box<dyn std::error::Error>> {
    let catalog = doc.catalog()?;
    let pages_ref = catalog.get(b"Pages")?.as_reference()?;
    Ok(pages_ref)
}

/// base 뒤에 add 문서의 모든 페이지를 append (Parent 교체 + Kids/Count 갱신)
fn append_doc(mut base: Document, mut add: Document) -> Result<Document, Box<dyn std::error::Error>> {
    let base_pages_id = pages_root_id(&base)?;
    let base_page_count = base.get_pages().len() as i64;

    add.renumber_objects_with(base.max_id + 1);

    let add_page_ids: Vec<ObjectId> = add.get_pages().values().cloned().collect();
    for pid in &add_page_ids {
        let obj = add.get_object_mut(*pid)?;
        let dict = obj.as_dict_mut()?;
        dict.set("Parent", base_pages_id);
    }

    base.objects.extend(add.objects);

    {
        let pages_obj = base.get_object_mut(base_pages_id)?;
        let pages_dict = pages_obj.as_dict_mut()?;
        let kids_obj = pages_dict.get_mut(b"Kids")?;
        let kids = kids_obj.as_array_mut()?;
        for pid in &add_page_ids {
            kids.push(Object::Reference(*pid));
        }
        let new_count = base_page_count + add_page_ids.len() as i64;
        pages_dict.set("Count", Object::Integer(new_count as i64));
    }

    base.renumber_objects();
    Ok(base)
}

/// 모든 페이지의 MediaBox/CropBox를 지정 크기로 통일
fn enforce_page_size(doc: &mut Document, w_pt: f64, h_pt: f64) -> Result<(), Box<dyn std::error::Error>> {
    let page_ids: Vec<ObjectId> = doc.get_pages().values().cloned().collect();
    let box_obj = Object::Array(vec![0.0.into(), 0.0.into(), w_pt.into(), h_pt.into()]);
    for pid in page_ids {
        let obj = doc.get_object_mut(pid)?;
        let dict = obj.as_dict_mut()?;
        dict.set("MediaBox", box_obj.clone());
        dict.set("CropBox",  box_obj.clone());
    }
    Ok(())
}

/// 지정 크기의 "빈 페이지 1장"만 가진 PDF 문서 생성
fn blank_page_doc(w_pt: f64, h_pt: f64) -> Document {
    let mut doc = Document::with_version("1.5");
    let pages_id = doc.new_object_id();
    let page_id = doc.new_object_id();
    let contents_id = doc.new_object_id();
    let catalog_id = doc.new_object_id();

    let stream = Stream::new(lopdf::Dictionary::new(), Vec::<u8>::new());
    doc.objects.insert(contents_id, Object::Stream(stream));

    let page_dict = dictionary! {
        "Type" => "Page",
        "Parent" => pages_id,
        "MediaBox" => Object::Array(vec![0.0.into(), 0.0.into(), w_pt.into(), h_pt.into()]),
        "CropBox"  => Object::Array(vec![0.0.into(), 0.0.into(), w_pt.into(), h_pt.into()]),
        "Resources" => lopdf::Dictionary::new(),
        "Contents" => contents_id,
    };
    doc.objects.insert(page_id, Object::Dictionary(page_dict));

    let pages_dict = dictionary! {
        "Type" => "Pages",
        "Kids" => Object::Array(vec![Object::Reference(page_id)]),
        "Count" => Object::Integer(1),
    };
    doc.objects.insert(pages_id, Object::Dictionary(pages_dict));

    let catalog = dictionary! {
        "Type" => "Catalog",
        "Pages" => pages_id,
    };
    doc.objects.insert(catalog_id, Object::Dictionary(catalog));
    doc.trailer.set(b"Root", catalog_id);
    doc
}

/// SVG → (벡터) **한 장짜리 페이지 PDF** 바이트 생성 (메모리)
///  - 페이지 크기: w_pt x h_pt
///  - 배치: **비율 유지(contain) + 중앙정렬**
fn svg_to_page_pdf_bytes(svg_path: &Path, w_pt: f64, h_pt: f64) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    // 1) SVG 파싱
    let svg_str = std::fs::read_to_string(svg_path)?;
    let mut opt = svg2pdf::usvg::Options::default();
    opt.fontdb_mut().load_system_fonts();
    let tree = svg2pdf::usvg::Tree::from_str(&svg_str, &opt)?;

    // 2) SVG → PDF Chunk(XObject) 변환
    let (svg_chunk, svg_root_ref) = svg2pdf::to_chunk(
        &tree,
        svg2pdf::ConversionOptions::default(),
    )
    .map_err(|e| {
        let err = std::io::Error::new(std::io::ErrorKind::Other, format!("svg2pdf to_chunk failed: {e}"));
        Box::<dyn std::error::Error>::from(err)
    })?; // returns (Chunk, Ref)
    // 참고: 공식 예제에서 이 Chunk를 페이지에 임베드하고 transform 행렬로 배치함. :contentReference[oaicite:2]{index=2}

    // 3) pdf-writer로 단일 페이지 구성 + transform 계산
    let mut alloc = PdfRef::new(1);
    let catalog_id   = alloc.bump();
    let page_tree_id = alloc.bump();
    let page_id      = alloc.bump();
    let content_id   = alloc.bump();
    let svg_name     = Name(b"S1");

    // chunk 리넘버링해서 우리 PDF ID 공간으로 편입
    let mut map = HashMap::new();
    let svg_chunk = svg_chunk.renumber(|old| *map.entry(old).or_insert_with(|| alloc.bump()));
    let svg_id = *map.get(&svg_root_ref).expect("svg root ref missing after renumber");

    // 페이지 생성
    let mut pdf = Pdf::new();
    pdf.catalog(catalog_id).pages(page_tree_id);
    pdf.pages(page_tree_id).kids([page_id]).count(1);

    // MediaBox 설정
    let mut page = pdf.page(page_id);
    page.media_box(Rect::new(0.0, 0.0, w_pt as f32, h_pt as f32));
    page.parent(page_tree_id);
    page.contents(content_id);

    // 리소스: XObject 등록
    let mut res = page.resources();
    res.x_objects().pair(svg_name, svg_id);
    res.finish();
    page.finish();

    // ===== 변환 행렬 계산 (contain + center) =====
    // XObject는 1pt × 1pt 단위를 기본으로 삼으므로,
    // 균등 스케일 s를 선택하고 가운데로 이동(tx, ty)시킵니다.
    // (이 방식은 SVG의 종횡비를 유지하고, 페이지 내에 letterbox가 생길 수 있음)
    let s = w_pt.min(h_pt);            // 사용 가능한 폭/높이 중 작은 값
    let tx = (w_pt - s) / 2.0;         // 중앙 정렬 X
    let ty = (h_pt - s) / 2.0;         // 중앙 정렬 Y

    let mut content = Content::new();
    content
        .transform([s as f32, 0.0, 0.0, s as f32, tx as f32, ty as f32])
        .x_object(svg_name);

    pdf.stream(content_id, &content.finish());
    // SVG 오브젝트 실제 바디 추가
    pdf.extend(&svg_chunk);

    Ok(pdf.finish())
}

fn make_pdf(args: Args, output: String) -> Result<(), Box<dyn std::error::Error>> {
    let unit = args.r#type.as_str();
    let w_pt = to_points(args.width, unit);
    let h_pt = to_points(args.height, unit);

    // 입력/출력 경로
    let front = PathBuf::from("./materials/front_matter.pdf");
    let back  = PathBuf::from("./materials/back_matter.pdf");
    let svgs_glob = "./materials/svg/*.svg";
    let out   = PathBuf::from(output);

    // temp (front/back 안정화용)
    let temp_dir = PathBuf::from("./temp");
    fs::create_dir_all(&temp_dir)?;
    let temp_front = temp_dir.join("front_matter.parsed.pdf");
    let temp_back  = temp_dir.join("back_matter.parsed.pdf");
    roundtrip_save(&front, &temp_front)?;
    roundtrip_save(&back,  &temp_back )?;

    // front 로드 + 페이지 크기 통일
    let mut merged = Document::load(&temp_front)?;
    enforce_page_size(&mut merged, w_pt, h_pt)?;

    // make-even: front가 홀수면 1장 추가
    if args.make_even {
        let front_pages = merged.get_pages().len();
        if front_pages % 2 == 1 {
            let blank = blank_page_doc(w_pt, h_pt);
            merged = append_doc(merged, blank)?;
        }
    }

    // SVG들: 메모리에서 **페이지 단위 PDF** 생성(변환 포함) → 병합
    let mut svg_paths: Vec<PathBuf> = glob(svgs_glob)?.filter_map(|e| e.ok()).collect();
    svg_paths.sort();

    // isARC=true → 사이 빈페이지 X, isARC=false → 사이 빈페이지 O
    let insert_between = !args.arc;

    for (i, svg) in svg_paths.iter().enumerate() {
        let svg_page_bytes = svg_to_page_pdf_bytes(svg, w_pt, h_pt)?;
        let svg_page_doc = Document::load_mem(&svg_page_bytes)?;
        merged = append_doc(merged, svg_page_doc)?;

        if insert_between && i + 1 <= svg_paths.len() {
            let blank = blank_page_doc(w_pt, h_pt);
            merged = append_doc(merged, blank)?;
        }
    }

    // back 로드 + 크기 통일 후 병합
    let mut back_doc = Document::load(&temp_back)?;
    enforce_page_size(&mut back_doc, w_pt, h_pt)?;
    merged = append_doc(merged, back_doc)?;

    // 최종 크기 통일(안전)
    enforce_page_size(&mut merged, w_pt, h_pt)?;

    if args.arc {
        process_pages::post_process_arc(&mut merged)?;
    }

    merged.save(out)?;
    println!("Done.");
    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    make_pdf(Args::arc(), String::from("./book_ARC.pdf"))?;
    make_pdf(Args::book(), String::from("./book.pdf"))?;
    Ok(())
}
