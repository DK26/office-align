use std::{error::Error, path::{Path, PathBuf}, str::FromStr};

// use office_align::high_zip::extract;

use walkdir::{DirEntry, WalkDir};
use zip::write::FileOptions;
use zip_extensions::{zip_create_from_directory_with_options, zip_extract};

fn is_xml(entry: &DirEntry) -> bool {
    entry.file_name()
         .to_str()
         .map(|s| s.ends_with(".xml"))
         .unwrap_or(false)
}

fn main() -> Result<(), Box<dyn Error>> {
    // Create a path:
    // let fname = std::path::Path::new("My Path");
    // extract(
    //     r"D:\git\office-align\sample\1.pptx",
    //     r"D:\git\office-align\sample\test_out",
    // )?;

    let archive_file = PathBuf::from_str(r"D:\git\office-align\sample\1.pptx").unwrap();
    let extraction_path =  PathBuf::from_str(r"D:\git\office-align\sample\test_out").unwrap();
    let new_archive = PathBuf::from_str(r"D:\git\office-align\sample\new_pptx.pptx").unwrap();

    // let archive_file = Path::new(r"D:\git\office-align\sample\1.pptx");
    // let extraction_path =  Path::new(r"D:\git\office-align\sample\test_out");
    // let new_archive = Path::new(r"D:\git\office-align\sample\new_pptx.pptx");

    // Extract the `.pptx` file
    zip_extract(&archive_file,&extraction_path)?;

    // TODO: Scan extracted directory slides
    let slides_dir = PathBuf::from_str(r"D:\git\office-align\sample\new_pptx.pptx").unwrap();

    let mut slides_path = extraction_path.clone();
    slides_path.push("ppt");
    slides_path.push("slides");

    let walker = WalkDir::new(&slides_path).max_depth(1);
    // for entry in walker.into_iter().filter_map(Result::ok).filter(|f|f.path().ends_with(".xml")) {
    //     println!("{}", entry.path().display());
    // }
    walker.into_iter().filter_map(Result::ok).filter(|f|f.path().ends_with(".xml")).for_each(|f|{
        println!("{}", f.path().display());
    });

    // TODO: Detect all hebrew letters
    // TODO: Replace with reversed version
    // TODO: Attempt to find a way to properly align right

    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // Recreate the `.pptx` file
    zip_create_from_directory_with_options(&new_archive, &extraction_path, options)?;

    Ok(())
}
