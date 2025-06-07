function handleFileUpload(event) {
    const files = event.target.files;
    if (files.length === 0) {
        console.log("No files selected.");
        return;
    }
    console.log("File selected:", files[0].name);
    fileInfo.textContent = `File selected: ${files[0].name}`;

    // Temporarily comment out the XLSX reading part
    // const fileReaders = Array.from(files).map(...);
    // Promise.all(fileReaders)...

    showMessage('File selected, but not processed for now (for debugging).', 'success');
    // You might want to remove columnMapping.style.display = 'flex'; if it was here
}
