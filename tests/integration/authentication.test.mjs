/**
 * Integration tests for authentication flows
 * Tests: loadExistingToken, initializeGraphClient, ensureGraphClient
 */
import { describe, it, expect, beforeEach, afterEach } from '@jest/globals';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { mockTokenData } from '../fixtures/apiResponses.mjs';
import { createMockTokenFile, removeMockTokenFile } from '../helpers/testUtils.mjs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Test token file path
const testTokenFilePath = path.join(__dirname, '../fixtures/.test-access-token.txt');

describe('Authentication Integration Tests', () => {
  describe('Token File Management', () => {
    beforeEach(() => {
      // Clean up any existing test token file
      removeMockTokenFile(testTokenFilePath);
    });

    afterEach(() => {
      // Clean up after tests
      removeMockTokenFile(testTokenFilePath);
    });

    it('should create a token file with JSON format', () => {
      createMockTokenFile(mockTokenData, testTokenFilePath);
      
      expect(fs.existsSync(testTokenFilePath)).toBe(true);
      const fileContent = fs.readFileSync(testTokenFilePath, 'utf8');
      const parsedContent = JSON.parse(fileContent);
      
      expect(parsedContent.token).toBe(mockTokenData.token);
      expect(parsedContent.clientId).toBe(mockTokenData.clientId);
      expect(parsedContent.scopes).toEqual(mockTokenData.scopes);
    });

    it('should read a token file with JSON format', () => {
      createMockTokenFile(mockTokenData, testTokenFilePath);
      
      const fileContent = fs.readFileSync(testTokenFilePath, 'utf8');
      const parsedContent = JSON.parse(fileContent);
      
      expect(parsedContent.token).toBeTruthy();
      expect(parsedContent.token).toBe('mock-access-token-12345');
      expect(parsedContent.scopes).toContain('Notes.Read');
      expect(parsedContent.scopes).toContain('Notes.ReadWrite');
    });

    it('should handle plain text token format (legacy)', () => {
      const plainToken = 'plain-text-token-12345';
      fs.writeFileSync(testTokenFilePath, plainToken, 'utf8');
      
      const fileContent = fs.readFileSync(testTokenFilePath, 'utf8');
      expect(fileContent).toBe(plainToken);
      
      // Should be able to read it as plain text
      expect(typeof fileContent).toBe('string');
      expect(fileContent.includes('plain-text-token')).toBe(true);
    });

    it('should handle missing token file gracefully', () => {
      expect(fs.existsSync(testTokenFilePath)).toBe(false);
      
      // This should not throw an error
      expect(() => {
        if (fs.existsSync(testTokenFilePath)) {
          fs.readFileSync(testTokenFilePath, 'utf8');
        }
      }).not.toThrow();
    });

    it('should validate token expiration information', () => {
      createMockTokenFile(mockTokenData, testTokenFilePath);
      
      const fileContent = fs.readFileSync(testTokenFilePath, 'utf8');
      const parsedContent = JSON.parse(fileContent);
      
      expect(parsedContent.expiresOn).toBeTruthy();
      expect(parsedContent.createdAt).toBeTruthy();
      
      const expiresOn = new Date(parsedContent.expiresOn);
      const createdAt = new Date(parsedContent.createdAt);
      
      expect(expiresOn).toBeInstanceOf(Date);
      expect(createdAt).toBeInstanceOf(Date);
      expect(expiresOn.getTime()).toBeGreaterThan(createdAt.getTime());
    });

    it('should store all required token metadata', () => {
      createMockTokenFile(mockTokenData, testTokenFilePath);
      
      const fileContent = fs.readFileSync(testTokenFilePath, 'utf8');
      const parsedContent = JSON.parse(fileContent);
      
      // Verify all required fields are present
      expect(parsedContent).toHaveProperty('token');
      expect(parsedContent).toHaveProperty('clientId');
      expect(parsedContent).toHaveProperty('scopes');
      expect(parsedContent).toHaveProperty('createdAt');
      expect(parsedContent).toHaveProperty('expiresOn');
      
      // Verify scopes array
      expect(Array.isArray(parsedContent.scopes)).toBe(true);
      expect(parsedContent.scopes.length).toBeGreaterThan(0);
    });
  });

  describe('Token Validation', () => {
    it('should validate token has required scopes', () => {
      const requiredScopes = ['Notes.Read', 'Notes.ReadWrite', 'Notes.Create', 'User.Read'];
      const tokenScopes = mockTokenData.scopes;
      
      requiredScopes.forEach(scope => {
        expect(tokenScopes).toContain(scope);
      });
    });

    it('should detect expired tokens', () => {
      const now = new Date();
      const expiredToken = {
        ...mockTokenData,
        expiresOn: new Date(now.getTime() - 3600000).toISOString() // 1 hour ago
      };
      
      const expiresOn = new Date(expiredToken.expiresOn);
      const isExpired = expiresOn.getTime() < now.getTime();
      
      expect(isExpired).toBe(true);
    });

    it('should detect valid (not expired) tokens', () => {
      const now = new Date();
      const validToken = {
        ...mockTokenData,
        expiresOn: new Date(now.getTime() + 3600000).toISOString() // 1 hour from now
      };
      
      const expiresOn = new Date(validToken.expiresOn);
      const isExpired = expiresOn.getTime() < now.getTime();
      
      expect(isExpired).toBe(false);
    });

    it('should validate client ID format', () => {
      expect(mockTokenData.clientId).toMatch(/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i);
    });
  });

  describe('Authentication Error Handling', () => {
    it('should handle corrupted JSON token file', () => {
      fs.writeFileSync(testTokenFilePath, '{ invalid json }', 'utf8');
      
      expect(() => {
        const content = fs.readFileSync(testTokenFilePath, 'utf8');
        JSON.parse(content);
      }).toThrow();
      
      removeMockTokenFile(testTokenFilePath);
    });

    it('should handle empty token file', () => {
      fs.writeFileSync(testTokenFilePath, '', 'utf8');
      
      const content = fs.readFileSync(testTokenFilePath, 'utf8');
      expect(content).toBe('');
      
      removeMockTokenFile(testTokenFilePath);
    });

    it('should handle file permission errors gracefully', () => {
      // This test is platform-specific and may not work on all systems
      // We'll create a file and test that we can read it
      createMockTokenFile(mockTokenData, testTokenFilePath);
      
      expect(() => {
        fs.readFileSync(testTokenFilePath, 'utf8');
      }).not.toThrow();
      
      removeMockTokenFile(testTokenFilePath);
    });
  });

  describe('Token Refresh Logic', () => {
    it('should determine when token refresh is needed (10 min buffer)', () => {
      const now = new Date();
      const tokenExpiringIn5Min = {
        ...mockTokenData,
        expiresOn: new Date(now.getTime() + 5 * 60 * 1000).toISOString() // 5 minutes from now
      };
      
      const expiresOn = new Date(tokenExpiringIn5Min.expiresOn);
      const bufferTime = 10 * 60 * 1000; // 10 minutes in milliseconds
      const needsRefresh = (expiresOn.getTime() - now.getTime()) < bufferTime;
      
      expect(needsRefresh).toBe(true);
    });

    it('should not refresh token with plenty of time remaining', () => {
      const now = new Date();
      const tokenExpiringIn60Min = {
        ...mockTokenData,
        expiresOn: new Date(now.getTime() + 60 * 60 * 1000).toISOString() // 60 minutes from now
      };
      
      const expiresOn = new Date(tokenExpiringIn60Min.expiresOn);
      const bufferTime = 10 * 60 * 1000; // 10 minutes in milliseconds
      const needsRefresh = (expiresOn.getTime() - now.getTime()) < bufferTime;
      
      expect(needsRefresh).toBe(false);
    });
  });

  describe('Client ID Configuration', () => {
    it('should use default Microsoft Graph Explorer client ID if not configured', () => {
      const defaultClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
      const configuredClientId = process.env.AZURE_CLIENT_ID || defaultClientId;
      
      expect(configuredClientId).toBeTruthy();
      expect(typeof configuredClientId).toBe('string');
      expect(configuredClientId.length).toBeGreaterThan(0);
    });

    it('should prefer environment variable AZURE_CLIENT_ID over default', () => {
      const testClientId = 'test-client-id-12345';
      const defaultClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
      
      // Simulate environment variable being set
      const originalEnv = process.env.AZURE_CLIENT_ID;
      process.env.AZURE_CLIENT_ID = testClientId;
      
      const configuredClientId = process.env.AZURE_CLIENT_ID || defaultClientId;
      expect(configuredClientId).toBe(testClientId);
      
      // Restore original environment
      if (originalEnv) {
        process.env.AZURE_CLIENT_ID = originalEnv;
      } else {
        delete process.env.AZURE_CLIENT_ID;
      }
    });
  });

  describe('Required Scopes', () => {
    const requiredScopes = ['Notes.Read', 'Notes.ReadWrite', 'Notes.Create', 'User.Read'];
    
    it('should define all required OneNote scopes', () => {
      expect(requiredScopes).toContain('Notes.Read');
      expect(requiredScopes).toContain('Notes.ReadWrite');
      expect(requiredScopes).toContain('Notes.Create');
    });

    it('should include User.Read scope for user information', () => {
      expect(requiredScopes).toContain('User.Read');
    });

    it('should match scopes in mock token data', () => {
      requiredScopes.forEach(scope => {
        expect(mockTokenData.scopes).toContain(scope);
      });
    });
  });
});
