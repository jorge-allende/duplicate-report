document.addEventListener("DOMContentLoaded", function () {
    const fileInput = document.getElementById("fileInput");
    const uploadButton = document.getElementById("uploadButton");

    // Disable the Upload button initially
    uploadButton.disabled = true;

    // Monitor the file input for changes
    fileInput.addEventListener("change", function () {
        if (fileInput.files.length > 0) {
            uploadButton.disabled = false; // Enable button
            uploadButton.classList.remove("disabled"); // Remove gray styling
        } else {
            uploadButton.disabled = true; // Disable button
            uploadButton.classList.add("disabled"); // Add gray styling
        }
    });

    // File Upload and Progress Handling
    uploadButton.addEventListener("click", function () {
        if (fileInput.files.length === 0) {
            alert("Please select a file.");
            return;
        }

        const file = fileInput.files[0];
        const formData = new FormData();
        formData.append("file", file);

        const xhr = new XMLHttpRequest();
        xhr.open("POST", "/upload", true);

        // Update progress bar
        xhr.upload.onprogress = function (event) {
            if (event.lengthComputable) {
                const percentComplete = (event.loaded / event.total) * 100;
                document.getElementById("progressBar").value = percentComplete;
                document.getElementById("progressText").textContent = Math.round(percentComplete) + "%";
            }
        };

        // On upload completion
        xhr.onload = function () {
            if (xhr.status === 200) {
                document.getElementById("downloadButton").style.display = "block";
                document.getElementById("downloadButton").disabled = false;
                document.getElementById("downloadButton").classList.add("enabled");
            } else {
                alert("Upload failed. Please try again.");
            }
        };

        xhr.send(formData);
    });

    // Handle Download Button Click
    document.getElementById("downloadButton").addEventListener("click", function () {
        window.location.href = "/download";
    });
});
