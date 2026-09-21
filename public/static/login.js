document.getElementById('loginForm').addEventListener('submit', async event => {
  event.preventDefault();
  const button = event.currentTarget.querySelector('button');
  const error = document.getElementById('loginError');
  button.disabled = true; error.textContent = '';
  try {
    const response = await fetch('/api/login', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({password:document.getElementById('password').value})});
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Sign-in failed.');
    location.assign('/');
  } catch (failure) { error.textContent = failure.message; }
  finally { button.disabled = false; }
});
