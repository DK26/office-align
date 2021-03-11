use std::{error::Error, fs::read_to_string, path::{Path, PathBuf}, str::FromStr};
use std::fs::File;

// use office_align::high_zip::extract;

use walkdir::{DirEntry, WalkDir};
use zip::write::FileOptions;
use zip_extensions::{zip_create_from_directory_with_options, zip_extract};

/// Returns `true` if a `DirEntry` is of given extension.
fn file_extension(entry: &DirEntry, extension: &str) -> bool {
    entry.file_name()
         .to_str()
         .map(|s| s.ends_with("." + extension))
         .unwrap_or(false)
}

fn reverse(text: &str) -> String {
    text.chars().rev().collect()
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

    // My Rust code is so ugly! Please tell me there is a better way
    let mut slides_path = extraction_path.clone();
    slides_path.push("ppt");
    slides_path.push("slides");

    let walker = WalkDir::new(&slides_path).max_depth(1);

    // for entry in walker.into_iter().filter_map(Result::ok).filter(|f|f.path().ends_with(".xml")) {
    //     println!("{}", entry.path().display());
    // }

    // walker.into_iter().filter_map(Result::ok).filter(|file|file.path().extension("xml")).for_each(|f|{
    //     println!("{}", f.path().display());
    // });

    // walker.into_iter().filter_map(Result::ok).filter(|f|f.path().ends_with(".xml")).for_each(|f|{
    //     println!("{}", f.path().display());
    // });

    // walker.into_iter().filter_map(Result::ok).filter(|file| is_xml(file)).for_each(|f|{
    //     println!("{}", f.path().display());
    // });

    fn str_strip_numbers(s: &str) -> Vec<i64> {
        lazy_static! {
            static ref RE: Regex = Regex::new(r"\d+").unwrap();
        }
        // iterate over all matches
        RE.find_iter(s)
            // try to parse the string matches as i64 (inferred from fn type signature)
            // and filter out the matches that can't be parsed (e.g. if there are too many digits to store in an i64).
            .filter_map(|digits| digits.as_str().parse().ok())
            // collect the results in to a Vec<i64> (inferred from fn type signature)
            .collect()
    }

    //walker.into_iter().filter_map(Result::ok).filter(|file| file_extension(file, ".xml")).for_each(|file_path|{
    for file_path in walker.into_iter().filter_map(Result::ok).filter(|file| file_extension(file, ".xml")) {
        
        println!("{}", file_path.path().display());
        
        let data = read_to_string(file_path.path()).expect("Unable to load contents");

        println!("data: {}", data);

        // // Open and read the file entirely
        // let mut open_file = File::open(&file_path.path());
        // match open_file {

        //     // Ignore error in opening file to not crash everything.
        //     // FIXME: Make it crash later because we don't want a in our context half-job
        //     Err(_) => continue,

        //     Ok(file) => {

        //         let mut data = String::new();

        //         //file.read_to_string(&mut data);
        //         file.read_to_string(&mut data).except("Unable to load file");

        //         drop(file);  // Close the file early

        //         println!("data: {}", data);

        //     }

        // }   
    };

    // for entry in glob(slides_path.join("*.xml")).expect("Failed to read glob pattern") {
    //     match entry {
    //         Ok(path) => println!("{:?}", path.display()),
    //         Err(e) => println!("{:?}", e),
    //     }
    // }

    // TODO: Detect all hebrew letters
    // TODO: Replace with reversed version
    // TODO: Attempt to find a way to properly align right

    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // Recreate the `.pptx` file
    zip_create_from_directory_with_options(&new_archive, &extraction_path, options)?;

    Ok(())
}
