// Global state
let currentClassId = localStorage.getItem('selectedClassId');
let classConfig = null;

// DOM Elements
const form = document.getElementById("checkinForm");
const msg = document.getElementById("msg");
const video = document.getElementById("video");
const canvas = document.getElementById("canvas");

// Load class configuration on startup
(async function loadClassConfig() {
  try {
    const response = await fetch('/config');
    classConfig = await response.json();
    
    if (!currentClassId && classConfig.defaultClass) {
      currentClassId = classConfig.defaultClass;
      localStorage.setItem('selectedClassId', currentClassId);
    }

    // Update page title and header
    updateClassDisplay();
  } catch (err) {
    console.error('Failed to load class config:', err);
  }
})();

// Update display when class changes
function updateClassDisplay() {
  if (!classConfig || !currentClassId) return;
  
  const selectedClass = classConfig.classes.find(c => c.id === currentClassId);
  if (selectedClass) {
    document.title = `${selectedClass.name} - Check-In`;
    document.querySelector('h2').textContent = `${selectedClass.name} Check-In`;
  }
}

// Camera initialization
(async function initCamera() {
  try {
    const stream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode: "user",
        width: { ideal: 4096 },  // Request large resolution
        height: { ideal: 2160 }
      },
      audio: false
    });
    video.srcObject = stream;
  } catch (err) {
    console.error("Camera error:", err);
    msg.textContent = "Camera access failed. Please allow camera permissions.";
  }
})();

function takeSnapshot() {
  const w = video.videoWidth;
  const h = video.videoHeight;
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(video, 0, 0, w, h);
  return canvas.toDataURL("image/png");
}

// Helper functions for styled messages
function showCheckinMessage(firstName) {
  msg.innerHTML = `
    <span class="checkmark">✓</span>
    <span class="greeting">Welcome, ${firstName}!</span>
    <span class="detail">You've been checked in successfully.</span>
  `;
  msg.className = 'msg success show';
}

function showAlreadyCheckedIn(firstName) {
  msg.innerHTML = `
    <span class="greeting">Hi, ${firstName}!</span>
    <span class="detail">You're already checked in today.</span>
  `;
  msg.className = 'msg info show';
}

function showErrorMessage(message) {
  msg.textContent = message;
  msg.className = 'msg error show';
}

function showLoadingMessage() {
  msg.innerHTML = `
    <span class="spinner"></span>
    <span class="detail">Checking in...</span>
  `;
  msg.className = 'msg info show';
}

// Form submission handler
form.addEventListener("submit", async (e) => {
  e.preventDefault();

  if (!currentClassId) {
    showErrorMessage("Please select a class period first");
    return;
  }

  const s_number = document.getElementById("snum").value.trim();
  if (!s_number) {
    showErrorMessage("Please enter your s-number");
    return;
  }

  showLoadingMessage();

  // Auto capture; no preview
  const image_data_url = takeSnapshot();

  try {
    const res = await fetch("/checkin", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ 
        s_number, 
        image_data_url,
        classId: currentClassId
      })
    });
    const data = await res.json();

    if (!data.ok) {
      showErrorMessage(data.error || "Check-in failed");
      return;
    }

    if (data.status === "already") {
      showAlreadyCheckedIn(data.first_name);
    } else {
      showCheckinMessage(data.first_name);
    }

    form.reset();
    document.getElementById("snum").focus();
  } catch (err) {
    console.error(err);
    showErrorMessage("Network error. Please try again");
  }
});
