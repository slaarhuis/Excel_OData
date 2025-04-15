import request from 'supertest';
import app from '../src/server';
import authService from '../src/services/auth.service';
import excelService from '../src/services/excel.service';

// Mock the auth service
jest.mock('../src/services/auth.service', () => ({
  getAccessToken: jest.fn().mockResolvedValue('mock-token'),
  validateBearerToken: jest.fn().mockImplementation((token) => token === 'valid-token')
}));

// Mock the excel service
jest.mock('../src/services/excel.service', () => ({
  getTableColumns: jest.fn().mockResolvedValue([
    { name: 'Column1' },
    { name: 'Column2' },
    { name: 'Column3' }
  ]),
  getTableRows: jest.fn().mockResolvedValue([
    { values: [['Value1', 'Value2', 'Value3']] },
    { values: [['Value4', 'Value5', 'Value6']] }
  ]),
  getTableRow: jest.fn().mockResolvedValue({ values: [['Value1', 'Value2', 'Value3']] }),
  convertToODataFormat: jest.fn().mockImplementation((rows, columns) => {
    return rows.map((row, index) => ({
      id: index.toString(),
      Column1: row.values[0][0],
      Column2: row.values[0][1],
      Column3: row.values[0][2]
    }));
  })
}));

describe('OData API Tests', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('Health endpoint should return 200', async () => {
    const response = await request(app).get('/health');
    expect(response.status).toBe(200);
    expect(response.body).toEqual({ status: 'ok' });
  });

  test('API info endpoint should return 200', async () => {
    const response = await request(app).get('/');
    expect(response.status).toBe(200);
    expect(response.body).toHaveProperty('name');
    expect(response.body).toHaveProperty('version');
    expect(response.body).toHaveProperty('endpoints');
  });

  test('OData endpoint should require authentication', async () => {
    const response = await request(app).get('/odata/ExcelRow');
    expect(response.status).toBe(401);
  });

  test('OData endpoint should reject invalid token', async () => {
    const response = await request(app)
      .get('/odata/ExcelRow')
      .set('Authorization', 'Bearer invalid-token');
    expect(response.status).toBe(403);
  });

  test('OData endpoint should return data with valid token', async () => {
    const response = await request(app)
      .get('/odata/ExcelRow')
      .set('Authorization', 'Bearer valid-token');
    expect(response.status).toBe(200);
    expect(response.body).toHaveProperty('value');
    expect(Array.isArray(response.body.value)).toBe(true);
  });

  test('OData metadata endpoint should return metadata', async () => {
    const response = await request(app)
      .get('/odata/$metadata')
      .set('Authorization', 'Bearer valid-token');
    expect(response.status).toBe(200);
    expect(response.header['content-type']).toContain('application/xml');
  });

  test('OData endpoint should return a single entity', async () => {
    const response = await request(app)
      .get('/odata/ExcelRow(\'0\')')
      .set('Authorization', 'Bearer valid-token');
    expect(response.status).toBe(200);
    expect(response.body).toHaveProperty('id', '0');
  });
});
