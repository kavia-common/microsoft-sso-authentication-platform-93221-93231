import React, { useEffect, useState } from 'react';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate, useAccount } from '@azure/msal-react';
import { InteractionType } from '@azure/msal-browser';
import axios from 'axios';
import './MicrosoftLogin.css';

const msalConfig = {
  auth: {
    clientId: process.env.REACT_APP_MSAL_CLIENT_ID || 'YOUR_CLIENT_ID_HERE', // Replace with actual Client ID
    authority: process.env.REACT_APP_MSAL_AUTHORITY || 'https://login.microsoftonline.com/common', // Or your tenant
    redirectUri: process.env.REACT_APP_MSAL_REDIRECT_URI || window.location.origin,
  },
};

const loginRequest = {
  scopes: ['openid', 'profile', 'email'],
};

// PUBLIC_INTERFACE
function MicrosoftLogin({ onSession, onError }) {
  /**
   * Centered dark-themed login form with prominent Microsoft SSO button.
   * Manages authentication, session, and error states.
   */

  const { instance, accounts, inProgress } = useMsal();
  const account = useAccount(accounts[0] || {});
  const [error, setError] = useState('');
  const [status, setStatus] = useState('init'); // 'init', 'loading', 'success', 'error'
  const [profile, setProfile] = useState(null);

  // Helper to initiate login popup
  const handleLogin = async () => {
    setStatus('loading');
    setError('');
    try {
      const loginResponse = await instance.loginPopup(loginRequest);
      setStatus('success');
      // Attempt session (REST call to backend)
      const accessToken = loginResponse.accessToken || (await instance.acquireTokenSilent(loginRequest)).accessToken;
      await authenticateWithBackend(accessToken);
    } catch (e) {
      setStatus('error');
      setError(e.message || 'Authentication failed');
      if (onError) onError(e);
    }
  };

  // Exchange MS token with backend Django for session (pseudo-endpoint /api/auth/microsoft/)
  const authenticateWithBackend = async (accessToken) => {
    try {
      const res = await axios.post(
        '/api/auth/microsoft/', // This endpoint should be implemented in Django backend
        { access_token: accessToken },
        { withCredentials: true }
      );
      setProfile(res.data.profile || null);
      if (onSession) onSession(res.data);
      setStatus('success');
    } catch (e) {
      setError('Backend authentication failed: ' + (e.response?.data?.detail || e.message));
      setStatus('error');
    }
  };

  // PUBLIC_INTERFACE
  const handleLogout = async () => {
    await instance.logoutPopup({ mainWindowRedirectUri: msalConfig.auth.redirectUri });
    setProfile(null);
    setStatus('init');
  };

  // Try auto-session if already signed in
  useEffect(() => {
    if (account && status === 'init') {
      (async () => {
        setStatus('loading');
        try {
          const { accessToken } = await instance.acquireTokenSilent(loginRequest);
          await authenticateWithBackend(accessToken);
        } catch {
          setStatus('init');
        }
      })();
    }
    // eslint-disable-next-line
  }, [account]);

  return (
    <div className="login-container">
      <div className="login-card">
        <img src="https://img.icons8.com/color/96/000000/microsoft.png" alt="Microsoft logo" className="ms-logo"/>
        <h2 className="login-title">Sign in with Microsoft</h2>
        <p className="login-desc">Use your Microsoft account to continue.</p>
        <UnauthenticatedTemplate>
          <button
            className="ms-login-btn"
            onClick={handleLogin}
            disabled={inProgress === 'login' || status === 'loading'}
            data-testid="ms-login-btn"
          >
            <span className="ms-login-btn-icon">
              <svg height="20" width="20" viewBox="0 0 23 23"><g><rect fill="#f35325" x="1" y="1" width="10" height="10"/><rect fill="#81bc06" x="12" y="1" width="10" height="10"/><rect fill="#05a6f0" x="1" y="12" width="10" height="10"/><rect fill="#ffba08" x="12" y="12" width="10" height="10"/></g></svg>
            </span>
            <span className="ms-login-btn-text">Sign in with Microsoft</span>
          </button>
        </UnauthenticatedTemplate>
        <AuthenticatedTemplate>
          <div className="login-success">
            <div className="login-welcome">
              <span role="img" aria-label="check" className="login-check">✔️</span>
              {profile ? (
                <span>
                  Welcome, <span className="login-user">{profile.displayName || profile.email || account?.name}</span>
                </span>
              ) : (
                <span>
                  Welcome, <span className="login-user">{account?.name || account?.username}</span>
                </span>
              )}
            </div>
            <button className="ms-logout-btn" onClick={handleLogout}>Logout</button>
          </div>
        </AuthenticatedTemplate>
        {status === 'loading' && (
          <div className="login-loading">Authenticating…</div>
        )}
        {!!error && (
          <div className="login-error" role="alert">
            {error}
          </div>
        )}
      </div>
      <footer className="login-footer">
        <small>Powered by Microsoft Identity • Secure SSO experience</small>
      </footer>
    </div>
  );
}

export default MicrosoftLogin;
