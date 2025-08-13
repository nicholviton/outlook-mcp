import { describe, it, expect, beforeEach } from 'vitest';
import { InputValidator, ValidationError } from '../../utils/InputValidator.js';

describe('InputValidator', () => {
  let validator;

  beforeEach(() => {
    validator = new InputValidator();
  });

  describe('Email Validation', () => {
    it('should validate correct email addresses', () => {
      const validEmails = [
        'test@example.com',
        'user+tag@domain.org',
        'firstname.lastname@company.co.uk',
        'user123@sub.domain.com'
      ];

      validEmails.forEach(email => {
        expect(validator.validateEmail(email)).toBe(true);
      });
    });

    it('should reject invalid email addresses', () => {
      const invalidEmails = [
        'invalid-email',
        'user@',
        '@domain.com',
        'user..name@domain.com',
        'user@domain',
        'user@domain..com',
        ''
      ];

      invalidEmails.forEach(email => {
        expect(validator.validateEmail(email)).toBe(false);
      });
    });

    it('should validate email arrays', () => {
      const validEmailArray = ['test1@example.com', 'test2@example.com'];
      const invalidEmailArray = ['test1@example.com', 'invalid-email'];

      expect(validator.validateEmailArray(validEmailArray)).toBe(true);
      expect(validator.validateEmailArray(invalidEmailArray)).toBe(false);
    });
  });

  describe('String Validation', () => {
    it('should validate strings within length limits', () => {
      expect(validator.validateString('test', 1, 10)).toBe(true);
      expect(validator.validateString('', 0, 10)).toBe(true);
      expect(validator.validateString('toolong', 1, 5)).toBe(false);
      expect(validator.validateString('', 1, 10)).toBe(false);
    });

    it('should sanitize strings', () => {
      const input = '<script>alert("xss")</script>Hello World';
      const sanitized = validator.sanitizeString(input);
      
      expect(sanitized).not.toContain('<script>');
      expect(sanitized).toContain('Hello World');
    });

    it('should validate HTML content', () => {
      const validHtml = '<p>Hello <strong>World</strong></p>';
      const invalidHtml = '<script>alert("xss")</script>';
      
      expect(validator.validateHtml(validHtml)).toBe(true);
      expect(validator.validateHtml(invalidHtml)).toBe(false);
    });
  });

  describe('Date Validation', () => {
    it('should validate ISO 8601 dates', () => {
      const validDates = [
        '2023-12-25T10:30:00Z',
        '2023-12-25T10:30:00+05:30',
        '2023-12-25T10:30:00.123Z',
        '2023-12-25'
      ];

      validDates.forEach(date => {
        expect(validator.validateDate(date)).toBe(true);
      });
    });

    it('should reject invalid dates', () => {
      const invalidDates = [
        'invalid-date',
        '2023-13-25T10:30:00Z',
        '2023-12-32T10:30:00Z',
        '2023-12-25T25:30:00Z',
        ''
      ];

      invalidDates.forEach(date => {
        expect(validator.validateDate(date)).toBe(false);
      });
    });

    it('should validate date ranges', () => {
      const start = '2023-12-25T10:00:00Z';
      const validEnd = '2023-12-25T11:00:00Z';
      const invalidEnd = '2023-12-25T09:00:00Z';

      expect(validator.validateDateRange(start, validEnd)).toBe(true);
      expect(validator.validateDateRange(start, invalidEnd)).toBe(false);
    });
  });

  describe('Tool Parameter Validation', () => {
    it('should validate email tool parameters', () => {
      const validParams = {
        to: ['test@example.com'],
        subject: 'Test Subject',
        body: 'Test Body',
        bodyType: 'text'
      };

      const invalidParams = {
        to: ['invalid-email'],
        subject: '',
        body: 'Test Body'
      };

      expect(() => validator.validateEmailParams(validParams)).not.toThrow();
      expect(() => validator.validateEmailParams(invalidParams)).toThrow(ValidationError);
    });

    it('should validate calendar event parameters', () => {
      const validParams = {
        subject: 'Test Event',
        start: {
          dateTime: '2023-12-25T10:00:00Z',
          timeZone: 'UTC'
        },
        end: {
          dateTime: '2023-12-25T11:00:00Z',
          timeZone: 'UTC'
        }
      };

      const invalidParams = {
        subject: '',
        start: {
          dateTime: 'invalid-date',
          timeZone: 'UTC'
        },
        end: {
          dateTime: '2023-12-25T09:00:00Z', // Earlier than start
          timeZone: 'UTC'
        }
      };

      expect(() => validator.validateEventParams(validParams)).not.toThrow();
      expect(() => validator.validateEventParams(invalidParams)).toThrow(ValidationError);
    });

    it('should validate search parameters', () => {
      const validParams = {
        query: 'test query',
        limit: 100,
        startDate: '2023-12-01T00:00:00Z',
        endDate: '2023-12-31T23:59:59Z'
      };

      const invalidParams = {
        query: '',
        limit: -1,
        startDate: '2023-12-31T23:59:59Z',
        endDate: '2023-12-01T00:00:00Z'
      };

      expect(() => validator.validateSearchParams(validParams)).not.toThrow();
      expect(() => validator.validateSearchParams(invalidParams)).toThrow(ValidationError);
    });
  });

  describe('Security Validation', () => {
    it('should detect malicious content', () => {
      const maliciousInputs = [
        '<script>alert("xss")</script>',
        'javascript:alert("xss")',
        'data:text/html,<script>alert("xss")</script>',
        'vbscript:msgbox("xss")',
        'onload=alert("xss")'
      ];

      maliciousInputs.forEach(input => {
        expect(validator.containsMaliciousContent(input)).toBe(true);
      });
    });

    it('should validate safe content', () => {
      const safeInputs = [
        'Normal text content',
        '<p>HTML paragraph</p>',
        '<strong>Bold text</strong>',
        'Email: test@example.com'
      ];

      safeInputs.forEach(input => {
        expect(validator.containsMaliciousContent(input)).toBe(false);
      });
    });

    it('should validate folder paths', () => {
      const validPaths = [
        'inbox',
        'sent',
        'drafts',
        'custom-folder',
        'parent/child-folder'
      ];

      const invalidPaths = [
        '../../../etc/passwd',
        '..\\..\\windows\\system32',
        'folder/../../../secret',
        'folder\\..\\..\\secret'
      ];

      validPaths.forEach(path => {
        expect(validator.validateFolderPath(path)).toBe(true);
      });

      invalidPaths.forEach(path => {
        expect(validator.validateFolderPath(path)).toBe(false);
      });
    });
  });

  describe('Schema Validation', () => {
    it('should validate objects against schemas', () => {
      const schema = {
        type: 'object',
        properties: {
          name: { type: 'string', minLength: 1 },
          age: { type: 'number', minimum: 0 },
          email: { type: 'string', format: 'email' }
        },
        required: ['name', 'email']
      };

      const validData = {
        name: 'John Doe',
        age: 30,
        email: 'john@example.com'
      };

      const invalidData = {
        name: '',
        age: -5,
        email: 'invalid-email'
      };

      expect(() => validator.validateSchema(validData, schema)).not.toThrow();
      expect(() => validator.validateSchema(invalidData, schema)).toThrow(ValidationError);
    });

    it('should validate arrays against schemas', () => {
      const schema = {
        type: 'array',
        items: { type: 'string', format: 'email' },
        minItems: 1,
        maxItems: 10
      };

      const validArray = ['test1@example.com', 'test2@example.com'];
      const invalidArray = ['test1@example.com', 'invalid-email'];
      const emptyArray = [];

      expect(() => validator.validateSchema(validArray, schema)).not.toThrow();
      expect(() => validator.validateSchema(invalidArray, schema)).toThrow(ValidationError);
      expect(() => validator.validateSchema(emptyArray, schema)).toThrow(ValidationError);
    });
  });

  describe('Business Logic Validation', () => {
    it('should validate attachment sizes', () => {
      const validSize = 10 * 1024 * 1024; // 10MB
      const invalidSize = 100 * 1024 * 1024; // 100MB

      expect(validator.validateAttachmentSize(validSize)).toBe(true);
      expect(validator.validateAttachmentSize(invalidSize)).toBe(false);
    });

    it('should validate file types', () => {
      const allowedTypes = ['image/jpeg', 'image/png', 'application/pdf'];
      const validType = 'image/jpeg';
      const invalidType = 'application/x-executable';

      expect(validator.validateFileType(validType, allowedTypes)).toBe(true);
      expect(validator.validateFileType(invalidType, allowedTypes)).toBe(false);
    });

    it('should validate recurrence patterns', () => {
      const validPattern = {
        type: 'daily',
        interval: 1
      };

      const invalidPattern = {
        type: 'invalid',
        interval: 0
      };

      expect(() => validator.validateRecurrencePattern(validPattern)).not.toThrow();
      expect(() => validator.validateRecurrencePattern(invalidPattern)).toThrow(ValidationError);
    });
  });

  describe('Batch Validation', () => {
    it('should validate multiple inputs', () => {
      const validInputs = [
        { type: 'email', value: 'test1@example.com' },
        { type: 'string', value: 'valid string', minLength: 1, maxLength: 100 },
        { type: 'date', value: '2023-12-25T10:00:00Z' }
      ];

      const invalidInputs = [
        { type: 'email', value: 'invalid-email' },
        { type: 'string', value: '', minLength: 1, maxLength: 100 },
        { type: 'date', value: 'invalid-date' }
      ];

      expect(() => validator.validateBatch(validInputs)).not.toThrow();
      expect(() => validator.validateBatch(invalidInputs)).toThrow(ValidationError);
    });

    it('should collect all validation errors', () => {
      const inputs = [
        { type: 'email', value: 'invalid-email' },
        { type: 'string', value: '', minLength: 1, maxLength: 100 }
      ];

      try {
        validator.validateBatch(inputs);
        expect.fail('Should have thrown validation error');
      } catch (error) {
        expect(error).toBeInstanceOf(ValidationError);
        expect(error.errors).toHaveLength(2);
      }
    });
  });
});

describe('ValidationError', () => {
  it('should create validation error with details', () => {
    const errors = [
      { field: 'email', message: 'Invalid email format' },
      { field: 'name', message: 'Name is required' }
    ];

    const error = new ValidationError('Validation failed', errors);

    expect(error.message).toBe('Validation failed');
    expect(error.errors).toEqual(errors);
    expect(error.name).toBe('ValidationError');
  });
});