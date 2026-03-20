// ============================
// Password Visibility Toggle
// ============================
function togglePasswordVisibility() {
  const passwordInput = document.getElementById('password');
  const eyeIcon = document.getElementById('eyeIcon');

  if (passwordInput.type === 'password') {
    passwordInput.type = 'text';
    eyeIcon.innerHTML = `
      <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94" />
      <path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19" />
      <line x1="1" y1="1" x2="23" y2="23" />
    `;
  } else {
    passwordInput.type = 'password';
    eyeIcon.innerHTML = `
      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
      <circle cx="12" cy="12" r="3" />
    `;
  }
}

// ============================
// Login Form Handler
// ============================
function handleLogin(event) {
  event.preventDefault();

  const username = document.getElementById('username');
  const password = document.getElementById('password');
  const loginBtn = document.getElementById('loginBtn');
  let isValid = true;

  // Reset states
  document.querySelectorAll('.input-wrapper').forEach(w => {
    w.classList.remove('error', 'success');
  });

  // Validate username
  if (!username.value.trim()) {
    username.closest('.input-wrapper').classList.add('error');
    document.getElementById('usernameGroup').classList.add('shake');
    setTimeout(() => document.getElementById('usernameGroup').classList.remove('shake'), 500);
    isValid = false;
  } else {
    username.closest('.input-wrapper').classList.add('success');
  }

  // Validate password
  if (!password.value.trim()) {
    password.closest('.input-wrapper').classList.add('error');
    document.getElementById('passwordGroup').classList.add('shake');
    setTimeout(() => document.getElementById('passwordGroup').classList.remove('shake'), 500);
    isValid = false;
  } else {
    password.closest('.input-wrapper').classList.add('success');
  }

  if (!isValid) return false;

  // Show loading state
  loginBtn.classList.add('loading');

  // Simulate login
  setTimeout(() => {
    loginBtn.classList.remove('loading');

    // Trigger character celebrations
    celebrateLogin();

    // Show success notification
    showNotification('🎉 登录成功！欢迎回来~');
  }, 1800);

  return false;
}

// ============================
// Celebration Effect
// ============================
function celebrateLogin() {
  const emojis = ['🎉', '✨', '🌟', '💫', '🎊', '💖', '🥳', '🪄'];
  const scene = document.querySelector('.characters-scene');
  const rect = scene.getBoundingClientRect();

  for (let i = 0; i < 12; i++) {
    setTimeout(() => {
      const emoji = document.createElement('span');
      emoji.className = 'emoji-burst';
      emoji.textContent = emojis[Math.floor(Math.random() * emojis.length)];
      emoji.style.left = (Math.random() * rect.width) + 'px';
      emoji.style.top = (Math.random() * rect.height * 0.5) + 'px';
      scene.appendChild(emoji);

      setTimeout(() => emoji.remove(), 1000);
    }, i * 80);
  }

  // Make characters do a happy bounce
  document.querySelectorAll('.character').forEach((char, i) => {
    setTimeout(() => {
      char.classList.add('clicked');
      setTimeout(() => char.classList.remove('clicked'), 500);
    }, i * 100);
  });
}

// ============================
// Notification Toast
// ============================
function showNotification(message) {
  // Remove any existing notification
  const existing = document.querySelector('.toast-notification');
  if (existing) existing.remove();

  const toast = document.createElement('div');
  toast.className = 'toast-notification';
  toast.innerHTML = message;
  toast.style.cssText = `
    position: fixed;
    top: 30px;
    left: 50%;
    transform: translateX(-50%) translateY(-20px);
    background: linear-gradient(135deg, #3A7BD5, #5B9CF5);
    color: white;
    padding: 14px 28px;
    border-radius: 14px;
    font-size: 15px;
    font-weight: 700;
    font-family: 'Nunito', sans-serif;
    box-shadow: 0 8px 30px rgba(59, 123, 213, 0.3);
    z-index: 1000;
    opacity: 0;
    transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
  `;

  document.body.appendChild(toast);

  requestAnimationFrame(() => {
    toast.style.opacity = '1';
    toast.style.transform = 'translateX(-50%) translateY(0)';
  });

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(-50%) translateY(-20px)';
    setTimeout(() => toast.remove(), 400);
  }, 3000);
}

// ============================
// Character Click Interactions
// ============================
document.addEventListener('DOMContentLoaded', () => {
  const characters = document.querySelectorAll('.character');
  const clickEmojis = {
    'char-1': ['👋', '💕', 'Hi~'],
    'char-2': ['🎈', '🎉', 'Yay!'],
    'char-3': ['😳', '💗', '///'],
    'char-4': ['🏃', '💨', 'Go!']
  };

  characters.forEach(char => {
    char.addEventListener('click', (e) => {
      char.classList.add('clicked');
      setTimeout(() => char.classList.remove('clicked'), 500);

      // Find which character class it has
      const charClass = Array.from(char.classList).find(c => c.startsWith('char-'));
      const emojis = clickEmojis[charClass] || ['✨'];

      // Create floating emoji/text
      const emoji = document.createElement('span');
      emoji.className = 'emoji-burst';
      emoji.textContent = emojis[Math.floor(Math.random() * emojis.length)];
      emoji.style.left = (e.offsetX || 40) + 'px';
      emoji.style.top = (e.offsetY || 20) + 'px';
      char.appendChild(emoji);

      setTimeout(() => emoji.remove(), 1000);
    });
  });

  // Add entrance animations stagger
  characters.forEach((char, i) => {
    char.style.opacity = '0';
    char.style.transform = 'translateY(30px)';
    setTimeout(() => {
      char.style.transition = 'opacity 0.5s ease, transform 0.5s ease';
      char.style.opacity = '1';
      char.style.transform = 'translateY(0)';
    }, 300 + i * 150);
  });

  // Input focus effects - make characters react
  const usernameInput = document.getElementById('username');
  const passwordInput = document.getElementById('password');

  usernameInput.addEventListener('focus', () => {
    document.querySelector('.char-1').style.animationDuration = '1.5s';
    document.querySelector('.char-2').style.animationDuration = '0.9s';
  });

  usernameInput.addEventListener('blur', () => {
    document.querySelector('.char-1').style.animationDuration = '3s';
    document.querySelector('.char-2').style.animationDuration = '1.8s';
  });

  passwordInput.addEventListener('focus', () => {
    // Shy character covers eyes
    document.querySelector('.char-3').style.animationDuration = '1.5s';
  });

  passwordInput.addEventListener('blur', () => {
    document.querySelector('.char-3').style.animationDuration = '4s';
  });
});
