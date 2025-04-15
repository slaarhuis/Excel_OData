import { Request, Response, NextFunction } from 'express';
import config from '../config/config';
import logger from '../utils/logger';

/**
 * Middleware to authenticate requests using Bearer token
 * This is required for Templafy integration which only supports static Bearer tokens
 */
export const authenticateToken = (req: Request, res: Response, next: NextFunction) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    logger.warn('Authentication failed: No token provided');
    return res.status(401).json({ error: 'Authentication required' });
  }

  // For Templafy, we use a static Bearer token as per their requirements
  if (token !== config.auth.bearerToken) {
    logger.warn('Authentication failed: Invalid token');
    return res.status(403).json({ error: 'Invalid token' });
  }

  next();
};
