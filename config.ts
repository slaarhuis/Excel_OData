import dotenv from 'dotenv';
import path from 'path';

// Load environment variables from .env file
dotenv.config({ path: path.resolve(__dirname, '../../.env') });

interface Config {
  port: number;
  nodeEnv: string;
  logLevel: string;
  sharepoint: {
    tenantId: string;
    clientId: string;
    clientSecret: string;
    siteUrl: string;
  };
  auth: {
    bearerToken: string;
    expiresIn: string;
  };
}

const config: Config = {
  port: parseInt(process.env.PORT || '3000', 10),
  nodeEnv: process.env.NODE_ENV || 'development',
  logLevel: process.env.LOG_LEVEL || 'info',
  sharepoint: {
    tenantId: process.env.SHAREPOINT_TENANT_ID || '',
    clientId: process.env.SHAREPOINT_CLIENT_ID || '',
    clientSecret: process.env.SHAREPOINT_CLIENT_SECRET || '',
    siteUrl: process.env.SHAREPOINT_SITE_URL || '',
  },
  auth: {
    bearerToken: process.env.API_BEARER_TOKEN || 'default-token-for-development-only',
    expiresIn: process.env.TOKEN_EXPIRES_IN || '1d',
  },
};

export default config;
