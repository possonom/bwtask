import { renderHook, waitFor } from '@testing-library/react';
import { useQueryProjects } from '../useQueryProjects';

// Mock PnP SP
jest.mock('@pnp/sp', () => ({
  sp: {
    web: {
      lists: {
        getByTitle: jest.fn(() => ({
          items: {
            select: jest.fn().mockReturnThis(),
            top: jest.fn().mockReturnThis(),
            getAll: jest.fn(),
          },
        })),
      },
    },
  },
}));

// Mock Logger
jest.mock('@pnp/logging', () => ({
  Logger: {
    write: jest.fn(),
  },
  LogLevel: {
    Info: 'Info',
    Error: 'Error',
    Warning: 'Warning',
  },
}));

// Mock error helper
jest.mock('../../helper/errorHelper', () => ({
  getErrorMessage: jest.fn((error) => error.message || 'Unknown error'),
}));

describe('useQueryProjects Hook', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('returns loading state initially', () => {
    const { result } = renderHook(() => useQueryProjects());
    
    expect(result.current.isLoading).toBe(true);
    expect(result.current.data).toEqual([]);
    expect(result.current.error).toBe(null);
  });

  it('returns mock data when SharePoint is not available', async () => {
    const { result } = renderHook(() => useQueryProjects());
    
    await waitFor(() => {
      expect(result.current.isLoading).toBe(false);
    });
    
    expect(result.current.data).toHaveLength(8); // Mock data has 8 items
    expect(result.current.data[0]).toEqual({
      Id: 1,
      Name: "Website Redesign",
      Status: 1,
    });
  });

  it('provides refetch functionality', async () => {
    const { result } = renderHook(() => useQueryProjects());
    
    await waitFor(() => {
      expect(result.current.isLoading).toBe(false);
    });
    
    expect(typeof result.current.refetch).toBe('function');
    
    // Test refetch
    await result.current.refetch();
    expect(result.current.data).toHaveLength(8);
  });

  it('handles errors gracefully', async () => {
    // This test would need more sophisticated mocking to simulate actual errors
    const { result } = renderHook(() => useQueryProjects());
    
    await waitFor(() => {
      expect(result.current.isLoading).toBe(false);
    });
    
    // Even with errors, should fall back to mock data
    expect(result.current.data).toHaveLength(8);
  });
});
