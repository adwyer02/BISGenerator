const dropzone = document.getElementById('uploadBox');
const fileInput = document.getElementById('upload');

dropzone.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
});

dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
});

dropzone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
});

dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');

    fileInput.files = e.dataTransfer.files;
    fileInput.dispatchEvent(new Event('change'));
});