import { describe, it, expect } from 'vitest';
import { createWorkbook } from '../src';

describe('Validation', () => {
  it('should validate sheet names', () => {
    const sf = createWorkbook();
    
    expect(() => sf.addSheet({
      name: '',
      headers: [],
      rows: []
    })).toThrow();

    expect(() => sf.addSheet({
      name: 'A'.repeat(32),
      headers: [],
      rows: []
    })).toThrow();

    expect(() => sf.addSheet({
      name: 'Invalid:',
      headers: [],
      rows: []
    })).toThrow();
  });

  it('should validate file path', async () => {
    const sf = createWorkbook();
    await expect(sf.save('')).rejects.toThrow();
  });
});
