document.addEventListener('DOMContentLoaded', (event) => {
  const fileInput = document.getElementById('file-input');
  const pdfDropBox = document.getElementById('pdfDropBox');
  const fileNameDisplay = document.getElementById('fileNameDisplay');

  const updateFileNameDisplay = (file) => {
    fileNameDisplay.textContent = `Selected file: ${file.name}`;
  };

  pdfDropBox.addEventListener('click', () => {
    fileInput.click(); // Open file dialog when the drop box is clicked
  });

  pdfDropBox.addEventListener('dragover', (e) => {
    e.preventDefault(); // Prevent default behavior when a file is dragged over the drop box
  });

  pdfDropBox.addEventListener('drop', (e) => {
    e.preventDefault(); // Prevent default behavior when a file is dropped
    if (e.dataTransfer.files.length) {
      fileInput.files = e.dataTransfer.files; // Update file input with the dropped files
      updateFileNameDisplay(e.dataTransfer.files[0]); // Update the file name display
    }
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length) {
      updateFileNameDisplay(fileInput.files[0]); // Update the file name display when a file is selected through the file dialog
    }
  });
});
