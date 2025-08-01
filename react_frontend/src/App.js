import React, { useEffect, useState } from 'react';
import { PublicClientApplication, EventType } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import MicrosoftLogin from './MicrosoftLogin';
import './App.css';

const msalConfig = {
  auth: {
    clientId: process.env.REACT_APP_MSAL_CLIENT_ID || 'YOUR_CLIENT_ID_HERE',
    authority: process.env.REACT_APP_MSAL_AUTHORITY || 'https://login.microsoftonline.com/common',
    redirectUri: process.env.REACT_APP_MSAL_REDIRECT_URI || window.location.origin,
  }
};
const msalInstance = new PublicClientApplication(msalConfig);

// PUBLIC_INTERFACE
function App() {
  /**
   * App entry renders a MSAL-wrapped Microsoft SSO login page with dark theme, session state, and error handling.
   * PublicClientApplication is shared via context.
   */
  const [theme, setTheme] = useState('dark');
  // Theme always dark, but allow toggle for future extensibility.
  useEffect(() => { document.documentElement.setAttribute('data-theme', theme); }, [theme]);

  // Session handling: optional, pass event handler if you want app-level behavior updates.
  const handleSession = (sessionObj) => {
    // Do something on session, e.g., change app state, store in Context/Redux etc.
    // This could be extended as needed.
    // console.log('Session:', sessionObj);
  };
  
  const handleError = (error) => {
    // Centralized error handler
    // console.error(error);
  };

  return (
    <MsalProvider instance={msalInstance}>
      <div className="App">
        <MicrosoftLogin onSession={handleSession} onError={handleError} />
      </div>
    </MsalProvider>
  );
}

export default App;
