document.getElementById('uploadButton').addEventListener('click', function() {
    const fileInput = document.getElementById('fileInput');
    if (fileInput.files.length === 0) {
        alert('Please select a file.');
        return;
    }

    const file = fileInput.files[0];
    const formData = new FormData();
    formData.append('file', file);

    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/upload', true);

    xhr.upload.onprogress = function(event) {
        if (event.lengthComputable) {
            const percentComplete = (event.loaded / event.total) * 100;
            document.getElementById('progressBar').value = percentComplete;
            document.getElementById('progressText').textContent = Math.round(percentComplete) + '%';
        }
    };

    xhr.onload = function() {
        if (xhr.status === 200) {
            document.getElementById('downloadButton').style.display = 'block';
            document.getElementById('downloadButton').disabled = false;
            document.getElementById('downloadButton').classList.add('enabled');
        } else {
            alert('Upload failed. Please try again.');
        }
    };

    xhr.send(formData);
});

document.getElementById('downloadButton').addEventListener('click', function() {
    window.location.href = '/download';
});
