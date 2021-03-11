//use std::fs::File;
use std::{collections::HashSet, error::Error, fs::{self, File, OpenOptions, read_to_string}, path::PathBuf, str::FromStr};
use std::io::prelude::*;  // This is a must for all `File` functionality to work

// use office_align::high_zip::extract;

use walkdir::{DirEntry, WalkDir};
use zip::write::FileOptions;
use zip_extensions::{zip_create_from_directory_with_options, zip_extract};

use lazy_static::lazy_static;
use regex::Regex;

/// Returns `true` if a `DirEntry` is of given extension.
fn file_extension(entry: &DirEntry, extension: &str) -> bool {
    entry
        .file_name()
        .to_str()
        .map(|s| s.ends_with(extension))
        .unwrap_or(false)
}

fn reverse(text: &str) -> String {
    text.chars().rev().collect()
}

fn extract_hebrew(s: &str) -> HashSet<&str> {
    // `lazy_static!` to run once and prevent loops from needlessly repeating this section
    lazy_static! {
        // static ref RE: Regex = Regex::new(r"<a:t>[\u0590-\u05FF]+\s?[\u0590-\u05FF]*</a:t>").unwrap();
        static ref RE: Regex = Regex::new(r#"[\u0590-\u05FF]+["']?[\u0590-\u05FF]*"#).unwrap();
    }

    let mut words_set = HashSet::new();
    for found_match in RE.find_iter(s) {
        words_set.insert(found_match.as_str());
    }
    // words_set.insert("test");
    // iterate over all matches
    // RE.find_iter(s).for_each(|m| words_set.insert(m.as_str()));  // WTF is the difference?
    words_set
}

fn main() -> Result<(), Box<dyn Error>> {
    // Create a path:
    // let fname = std::path::Path::new("My Path");
    // extract(
    //     r"D:\git\office-align\sample\1.pptx",
    //     r"D:\git\office-align\sample\test_out",
    // )?;

    let archive_file = PathBuf::from_str(r"D:\git\office-align\sample\1.pptx").unwrap();
    let extraction_path = PathBuf::from_str(r"D:\git\office-align\sample\test_out").unwrap();
    let new_archive = PathBuf::from_str(r"D:\git\office-align\sample\new_pptx.pptx").unwrap();

    // let archive_file = Path::new(r"D:\git\office-align\sample\1.pptx");
    // let extraction_path =  Path::new(r"D:\git\office-align\sample\test_out");
    // let new_archive = Path::new(r"D:\git\office-align\sample\new_pptx.pptx");

    println!("PreZip");
    // Extract the `.pptx` file
    zip_extract(&archive_file, &extraction_path)?;
    println!("PostZip");

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

    // fn str_strip_numbers(s: &str) -> Vec<i64> {
    //     lazy_static! {
    //         static ref RE: Regex = Regex::new(r"\d+").unwrap();
    //     }
    //     // iterate over all matches
    //     RE.find_iter(s)
    //         // try to parse the string matches as i64 (inferred from fn type signature)
    //         // and filter out the matches that can't be parsed (e.g. if there are too many digits to store in an i64).
    //         .filter_map(|digits| digits.as_str().parse().ok())
    //         // collect the results in to a Vec<i64> (inferred from fn type signature)
    //         .collect()
    // }

    //walker.into_iter().filter_map(Result::ok).filter(|file| file_extension(file, ".xml")).for_each(|file_path|{
    for file_path in walker
        .into_iter()
        .filter_map(Result::ok) // Return only the `Result::ok` variants
        .filter(|file| file_extension(file, ".xml"))
    {
        println!("{}", file_path.path().display());

        let data = read_to_string(file_path.path()).expect("Unable to open slide");

        //println!("data: {}", data);
        let hebrew_words = extract_hebrew(&data);

        let mut data = data.clone();
        for word in hebrew_words {
            //println!("File: {} Word: {}", file_path.path().display(), reverse(word));
            data = data.replace(word, &reverse(word));
        }

        let mut file = OpenOptions::new()
            .write(true)
            .truncate(true)
            .open(file_path.path())
            .expect("Unable to open slide");
        //let f:File;
        //f = file;
        file
            .write(data.as_bytes())
            .expect("Unable to write slide");
         
        drop(file);

        // println!("{}", data);
        // for mat in Regex::new(r"\w").unwrap().find_iter(&data) {
        //     println!("{:?}", mat);
        // }

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
    }

    // for entry in glob(slides_path.join("*.xml")).expect("Failed to read glob pattern") {
    //     match entry {
    //         Ok(path) => println!("{:?}", path.display()),
    //         Err(e) => println!("{:?}", e),
    //     }
    // }

    // TODO: Detect all hebrew letters
    // TODO: Replace with reversed version
    // TODO: Attempt to find a way to properly align right

    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // Recreate the `.pptx` file
    zip_create_from_directory_with_options(&new_archive, &extraction_path, options)?;

    // Clean Extraction
    fs::remove_dir_all(&extraction_path).expect("Unable to remove the extraction directory");

    Ok(())
}
