import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getAuth, signInWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";
import { firebaseConfig } from "../config/firebase-config.js";

// Firebase ක්‍රියාත්මක කිරීම
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);

// Login බොත්තම එබූ විට
document.getElementById('loginBtn').addEventListener('click', () => {
    const email = document.getElementById('email').value;
    const password = document.getElementById('password').value;
    const errorMsg = document.getElementById('errorMsg');
    const loginBtn = document.getElementById('loginBtn');

    if(!email || !password) {
        errorMsg.innerText = "කරුණාකර ඊමේල් සහ මුරපදය ඇතුළත් කරන්න.";
        errorMsg.classList.remove('d-none');
        return;
    }

    // බොත්තම එබූ පසු Loading වන බව පෙන්වීම (Modern විදිහට)
    loginBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> ලොග් වෙමින් පවතී...';
    loginBtn.disabled = true;
    errorMsg.classList.add('d-none');

    // Firebase හරහා ලොග් වීම
    signInWithEmailAndPassword(auth, email, password)
        .then((userCredential) => {
            window.location.href = "dashboard.html";
        })
        .catch((error) => {
            errorMsg.innerText = "ලොග් වීමේ දෝෂයක්: ඊමේල් හෝ මුරපදය වැරදියි.";
            errorMsg.classList.remove('d-none');
            // බොත්තම නැවත සාමාන්‍ය තත්ත්වයට පත් කිරීම
            loginBtn.innerHTML = '<i class="bi bi-arrow-right-circle-fill me-2" aria-hidden="true"></i>Login';
            loginBtn.disabled = false;
        });
});