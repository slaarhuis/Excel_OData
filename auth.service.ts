import axios from 'axios';
import qs from 'querystring';
import config from '../config/config';
import logger from '../utils/logger';

/**
 * Service for handling SharePoint authentication
 */
class AuthService {
  private accessToken: string | null = null;
  private tokenExpiration: Date | null = null;

  /**
   * Get access token for Microsoft Graph API
   * Uses client credentials flow for service-to-service authentication
   */
  async getAccessToken(): Promise<string> {
    try {
      // Check if we have a valid token already
      if (this.accessToken && this.tokenExpiration && this.tokenExpiration > new Date()) {
        logger.debug('Using cached access token');
        return this.accessToken;
      }

      const { tenantId, clientId, clientSecret } = config.sharepoint;
      
      // Validate required configuration
      if (!tenantId || !clientId || !clientSecret) {
        throw new Error('Missing SharePoint authentication configuration');
      }

      logger.info('Requesting new access token from Microsoft Identity platform');
      
      // Token endpoint for Microsoft Identity platform
      const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
      
      // Request body for client credentials flow
      const requestBody = {
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      };

      // Make request to token endpoint
      const response = await axios.post(
        tokenEndpoint,
        qs.stringify(requestBody),
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
        }
      );

      // Extract token and expiration
      const { access_token, expires_in } = response.data;
      
      // Calculate expiration date (subtract 5 minutes for safety margin)
      const expirationDate = new Date();
      expirationDate.setSeconds(expirationDate.getSeconds() + expires_in - 300);
      
      // Store token and expiration
      this.accessToken = access_token;
      this.tokenExpiration = expirationDate;
      
      logger.info('Successfully obtained access token', { 
        expiresAt: this.tokenExpiration.toISOString() 
      });
      
      return access_token;
    } catch (error: any) {
      logger.error('Failed to get access token', { 
        error: error.message,
        response: error.response?.data 
      });
      throw new Error('SharePoint authentication failed');
    }
  }

  /**
   * Validate Templafy Bearer token
   */
  validateBearerToken(token: string): boolean {
    return token === config.auth.bearerToken;
  }
}

export default new AuthService();
