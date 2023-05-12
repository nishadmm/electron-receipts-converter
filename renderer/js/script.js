const form = document.querySelector("#img-form");
const img = document.querySelector("#img");
const outputPath = document.querySelector("#output-path");
const filename = document.querySelector("#filename");
const heightInput = document.querySelector("#height");
const widthInput = document.querySelector("#width");

// Resize image
function resizeImage(e) {
  e.preventDefault();

  if (!img.files[0]) {
    alertError("Please upload an receipt");
    return;
  }

  // Electron adds a bunch of extra properties to the file object including the path
  const imgPath = img.files[0].path;

  // Send to main using ipcRenderer
  ipcRenderer.send("receipt:convert", {
    imgPath,
    filesPath: [...img.files].map((singleFile) => singleFile.path),
  });
}

// When done, show message
ipcRenderer.on("receipt:done", () =>
  alertSuccess(`Receipt converted successfully`)
);

function loadImage(e) {
  const file = e.target.files[0];

  console.log(e.target.files[0].originalname);

  const files = img.files;

  for (let i = 0; i < files.length; i++) {
    const fileItem = files[i];

    if (!isFileImage(fileItem)) {
      alertError("Please upload all receipts in PDF format.");
      return false;
    }
  }

  form.style.display = "block";
  filename.innerText = file?.name;
  outputPath.innerText = path.join(os.homedir(), "receipt-converter");
}

// Make sure file is image
function isFileImage(file) {
  const acceptedImageTypes = ["application/pdf"];

  return file && acceptedImageTypes.includes(file["type"]);
}

function alertSuccess(message) {
  Toastify.toast({
    text: message,
    duration: 5000,
    close: false,
    style: {
      background: "green",
      color: "white",
      textAlign: "center",
    },
  });
}

function alertError(message) {
  Toastify.toast({
    text: message,
    duration: 5000,
    close: false,
    style: {
      background: "red",
      color: "white",
      textAlign: "center",
    },
  });
}

// File select listener
img.addEventListener("change", loadImage);
// Form submit listener
form.addEventListener("submit", resizeImage);
