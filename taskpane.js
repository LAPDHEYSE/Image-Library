const enableLogging = false; // Set to false to disable logging

function log(message) {
  if (enableLogging) {
    console.log(message);
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    log("Office is ready.");
    run();
  } else {
    console.error("This add-in is not running in PowerPoint.");
  }
});

async function run() {
  try {
    const container = document.getElementById("container");
    if (!container) {
      console.error("Container element not found.");
      return;
    }

    container.innerHTML = ""; // Clear any existing content

    let imageLibrary = JSON.parse(localStorage.getItem("imageLibrary")) || [];

    function saveImageLibrary() {
      try {
        localStorage.setItem("imageLibrary", JSON.stringify(imageLibrary));
      } catch (e) {
        if (e.name === 'QuotaExceededError' || e.name === 'NS_ERROR_DOM_QUOTA_REACHED') {
          alert("Storage limit exceeded. Please delete some images.");
        } else {
          console.error("Error saving to local storage:", e);
        }
      }
    }

    function addImageToLibrary(name, url) {
      const image = { name, url };
      imageLibrary.push(image);
      saveImageLibrary();
      displayImage(image);
    }

    function displayImage(image) {
      const imageContainer = document.createElement("div");
      imageContainer.className = "image-container";

      const imageIcon = document.createElement("img");
      imageIcon.src = image.url;
      imageIcon.width = 50;
      imageIcon.height = 50;
      imageIcon.alt = image.name;

      imageIcon.onclick = () => {
        try {
          const base64data = image.url.split(',')[1];
          Office.context.document.setSelectedDataAsync(base64data, { coercionType: Office.CoercionType.Image }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.error("Error inserting image:", asyncResult.error.message);
            } else {
              log("Image inserted successfully.");
            }
          });
        } catch (error) {
          console.error("Error inserting image:", error);
        }
      };

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "checkbox";

      const filename = document.createElement("span");
      filename.className = "filename";
      filename.textContent = image.name.length > 7 ? image.name.substring(0, 7) + "..." : image.name;

      imageContainer.appendChild(imageIcon);
      imageContainer.appendChild(checkbox);
      imageContainer.appendChild(filename);
      container.appendChild(imageContainer);
    }

    // Display existing images from local storage
    imageLibrary.forEach(displayImage);

    // Handle drag-and-drop
    const dropArea = document.getElementById("drop-area");

    dropArea.addEventListener("dragover", (event) => {
      event.preventDefault();
      dropArea.classList.add("dragging");
    });

    dropArea.addEventListener("dragleave", () => {
      dropArea.classList.remove("dragging");
    });

    dropArea.addEventListener("drop", (event) => {
      event.preventDefault();
      dropArea.classList.remove("dragging");
      const files = event.dataTransfer.files;
      handleFiles(files);
    });

    function handleFiles(files) {
      for (const file of files) {
        const reader = new FileReader();
        reader.onload = (event) => {
          addImageToLibrary(file.name, event.target.result);
        };
        reader.readAsDataURL(file);
      }
    }

    // Handle URL import
    const importButton = document.getElementById("importButton");
    importButton.addEventListener("click", () => {
      const imageUrl = document.getElementById("imageUrl").value;
      if (imageUrl) {
        fetch(imageUrl)
          .then(response => response.blob())
          .then(blob => {
            const reader = new FileReader();
            reader.onload = (event) => {
              addImageToLibrary("Imported Image", event.target.result);
            };
            reader.readAsDataURL(blob);
          })
          .catch(error => {
            console.error("Error fetching image:", error);
          });
      }
    });

    // Handle delete selected images
    const deleteSelectedButton = document.getElementById("deleteSelectedButton");
    deleteSelectedButton.addEventListener("click", () => {
      const checkboxes = document.querySelectorAll(".checkbox");
      checkboxes.forEach((checkbox, index) => {
        if (checkbox.checked) {
          const imageContainer = checkbox.parentElement;
          container.removeChild(imageContainer);
          imageLibrary.splice(index, 1);
        }
      });
      saveImageLibrary();
    });

    // Timer functionality
    let timerInterval;
    let timeRemaining;
    let totalTime;
    let startTime;

    document.getElementById("startTimerButton").addEventListener("click", startTimer);
    document.getElementById("resetTimerButton").addEventListener("click", resetTimer);

    function startTimer() {
      const minutes = parseInt(document.getElementById("minutes").value) || 0;
      const seconds = parseInt(document.getElementById("seconds").value) || 0;
      totalTime = (minutes * 60) + seconds;
      timeRemaining = totalTime;
      startTime = Date.now();

      const timerDisplay = document.getElementById("timerDisplay");
      const timerProgress = document.getElementById("timerProgress");

      timerInterval = setInterval(() => {
        const elapsedTime = Math.floor((Date.now() - startTime) / 1000);
        timeRemaining = totalTime - elapsedTime;

        if (timeRemaining <= 0) {
          clearInterval(timerInterval);
          timerDisplay.textContent = "00:00";
          timerProgress.style.width = "0%";
          return;
        }

        const minutesRemaining = Math.floor(timeRemaining / 60);
        const secondsRemaining = timeRemaining % 60;
        timerDisplay.textContent = `${String(minutesRemaining).padStart(2, '0')}:${String(secondsRemaining).padStart(2, '0')}`;
        timerProgress.style.width = `${(timeRemaining / totalTime) * 100}%`;
      }, 1000);
    }

    function resetTimer() {
      clearInterval(timerInterval);
      timeRemaining = totalTime;
      const timerDisplay = document.getElementById("timerDisplay");
      const timerProgress = document.getElementById("timerProgress");
      const minutesRemaining = Math.floor(timeRemaining / 60);
      const secondsRemaining = timeRemaining % 60;
      timerDisplay.textContent = `${String(minutesRemaining).padStart(2, '0')}:${String(secondsRemaining).padStart(2, '0')}`;
      timerProgress.style.width = "100%";
    }

    log("UI elements added to container.");
  } catch (error) {
    console.error("Error running PowerPoint context:", error);
  }
}