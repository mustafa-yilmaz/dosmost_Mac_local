import { useState, useEffect } from 'react';
import { PublicClientApplication, AccountInfo } from '@azure/msal-browser';
import { msalConfig, loginRequest } from '../config/msal-config';
import { office365Service } from '../services/office365-service';

const msalInstance = new PublicClientApplication(msalConfig);

export const useMicrosoftAuth = () => {
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    // Initialize MSAL
    msalInstance.initialize().then(() => {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        setAccount(accounts[0]);
        acquireToken(accounts[0]);
      }
      setIsLoading(false);
    });
  }, []);

  const acquireToken = async (account: AccountInfo) => {
    try {
      const response = await msalInstance.acquireTokenSilent({
        ...loginRequest,
        account,
      });
      setAccessToken(response.accessToken);
      office365Service.setAccessToken(response.accessToken);
      return response.accessToken;
    } catch (error) {
      console.error('Silent token acquisition failed:', error);
      // Fallback to interactive method
      try {
        const response = await msalInstance.acquireTokenPopup(loginRequest);
        setAccessToken(response.accessToken);
        office365Service.setAccessToken(response.accessToken);
        return response.accessToken;
      } catch (popupError) {
        console.error('Popup token acquisition failed:', popupError);
        setError('Failed to acquire access token');
        return null;
      }
    }
  };

  const login = async () => {
    setIsLoading(true);
    setError(null);
    try {
      const response = await msalInstance.loginPopup(loginRequest);
      setAccount(response.account);
      setAccessToken(response.accessToken);
      office365Service.setAccessToken(response.accessToken);
    } catch (error: any) {
      setError(error.message || 'Login failed');
      console.error('Login error:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const logout = async () => {
    try {
      await msalInstance.logoutPopup();
      setAccount(null);
      setAccessToken(null);
      office365Service.setAccessToken('');
    } catch (error) {
      console.error('Logout error:', error);
    }
  };

  return {
    account,
    accessToken,
    isAuthenticated: !!account && !!accessToken,
    isLoading,
    error,
    login,
    logout,
  };
};
