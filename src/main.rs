use std::{error::Error, path::PathBuf, str::FromStr};

use office_align::high_zip::extract;
use zip::write::FileOptions;
use zip_extensions::{zip_create_from_directory, zip_create_from_directory_with_options, zip_extract};

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

    zip_extract(&archive_file,&extraction_path)?;

    // TODO: Detect all hebrew letters
    // TODO: Replace with reversed version
    // TODO: Attempt to find a way to proprly align right

    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip_create_from_directory_with_options(&new_archive, &extraction_path, options)?;

    Ok(())
}
