/*!
High level abstractions for the `zip` crate.
*/

use std::{fs, path::Path};

use zip::{result::ZipResult, write::FileOptions, CompressionMethod};

pub fn extract<P: AsRef<Path>>(archive_path: P, output_path: P) -> ZipResult<()> {

    let file = fs::File::open(&archive_path).expect("Unable to open archive file.");

    fs::create_dir_all(&output_path).expect("Unable to create output directories.");

    let mut archive = zip::ZipArchive::new(file).unwrap();

    archive.extract(output_path)?;

    Ok(())
}

pub fn archive<P: AsRef<Path>>(
    archive_file: P,
    output_path: P,
    compression_method: CompressionMethod,
) -> ZipResult<()> {
    let file = std::fs::File::create(&output_path).unwrap();

    let mut zip = zip::ZipWriter::new(file);

    let options = FileOptions::default()
        .compression_method(compression_method);

    zip.start_file("test/☃.txt", options)?;

    Ok(())
}
