/**
 * Tests for Zustand project store
 */

import { act, renderHook } from '@testing-library/react';
import { useProjectStore, useProjectSelectors, useProjectActions } from '../projectStore';
import { Project } from '../../interfaces/Project';
import { Status } from '../../interfaces/Status';

// Mock project data
const mockProjects: Project[] = [
  {
    Id: 1,
    Name: 'Test Project 1',
    Status: Status.InProgress,
    Priority: 1,
    Progress: 50,
  },
  {
    Id: 2,
    Name: 'Test Project 2',
    Status: Status.Completed,
    Priority: 2,
    Progress: 100,
  },
];

describe('Project Store', () => {
  beforeEach(() => {
    // Reset store state before each test
    useProjectStore.setState({
      projects: [],
      selectedProject: null,
      projectSummary: null,
      isLoading: false,
      error: null,
      filters: {},
      searchText: '',
      sortBy: 'name',
      sortDirection: 'asc',
      isPanelOpen: false,
      isCreateModalOpen: false,
      isEditMode: false,
    });
  });

  describe('Basic State Management', () => {
    it('should initialize with empty state', () => {
      const { result } = renderHook(() => useProjectStore());
      
      expect(result.current.projects).toEqual([]);
      expect(result.current.selectedProject).toBeNull();
      expect(result.current.isLoading).toBe(false);
      expect(result.current.error).toBeNull();
    });

    it('should set projects', () => {
      const { result } = renderHook(() => useProjectStore());
      
      act(() => {
        result.current.setProjects(mockProjects);
      });
      
      expect(result.current.projects).toEqual(mockProjects);
    });

    it('should add a project', () => {
      const { result } = renderHook(() => useProjectStore());
      const newProject: Project = {
        Id: 3,
        Name: 'New Project',
        Status: Status.NotStarted,
      };
      
      act(() => {
        result.current.setProjects(mockProjects);
        result.current.addProject(newProject);
      });
      
      expect(result.current.projects).toHaveLength(3);
      expect(result.current.projects[2]).toEqual(newProject);
    });

    it('should update a project', () => {
      const { result } = renderHook(() => useProjectStore());
      const updatedProject: Project = {
        ...mockProjects[0],
        Name: 'Updated Project Name',
        Status: Status.Completed,
      };
      
      act(() => {
        result.current.setProjects(mockProjects);
        result.current.updateProject(updatedProject);
      });
      
      expect(result.current.projects[0]).toEqual(updatedProject);
    });

    it('should remove a project', () => {
      const { result } = renderHook(() => useProjectStore());
      
      act(() => {
        result.current.setProjects(mockProjects);
        result.current.removeProject(1);
      });
      
      expect(result.current.projects).toHaveLength(1);
      expect(result.current.projects[0].Id).toBe(2);
    });
  });

  describe('Filtering and Sorting', () => {
    it('should filter projects by search text', () => {
      const { result } = renderHook(() => useProjectStore());
      
      act(() => {
        result.current.setProjects(mockProjects);
        result.current.setSearchText('Test Project 1');
      });
      
      expect(result.current.filteredProjects).toHaveLength(1);
      expect(result.current.filteredProjects[0].Name).toBe('Test Project 1');
    });

    it('should filter projects by status', () => {
      const { result } = renderHook(() => useProjectStore());
      
      act(() => {
        result.current.setProjects(mockProjects);
        result.current.setFilters({ status: Status.Completed });
      });
      
      expect(result.current.filteredProjects).toHaveLength(1);
      expect(result.current.filteredProjects[0].Status).toBe(Status.Completed);
    });

    it('should sort projects by name', () => {
      const { result } = renderHook(() => useProjectStore());
      
      act(() => {
        result.current.setProjects(mockProjects);
        result.current.setSorting('name', 'desc');
      });
      
      const sorted = result.current.sortedProjects;
      expect(sorted[0].Name).toBe('Test Project 2');
      expect(sorted[1].Name).toBe('Test Project 1');
    });
  });

  describe('UI State Management', () => {
    it('should manage panel state', () => {
      const { result } = renderHook(() => useProjectStore());
      
      act(() => {
        result.current.openPanel(mockProjects[0]);
      });
      
      expect(result.current.isPanelOpen).toBe(true);
      expect(result.current.selectedProject).toEqual(mockProjects[0]);
      expect(result.current.isEditMode).toBe(true);
      
      act(() => {
        result.current.closePanel();
      });
      
      expect(result.current.isPanelOpen).toBe(false);
      expect(result.current.selectedProject).toBeNull();
      expect(result.current.isEditMode).toBe(false);
    });

    it('should manage loading state', () => {
      const { result } = renderHook(() => useProjectStore());
      
      act(() => {
        result.current.setLoading(true);
      });
      
      expect(result.current.isLoading).toBe(true);
      
      act(() => {
        result.current.setLoading(false);
      });
      
      expect(result.current.isLoading).toBe(false);
    });

    it('should manage error state', () => {
      const { result } = renderHook(() => useProjectStore());
      const errorMessage = 'Test error message';
      
      act(() => {
        result.current.setError(errorMessage);
      });
      
      expect(result.current.error).toBe(errorMessage);
      expect(result.current.isLoading).toBe(false);
    });
  });

  describe('Selectors', () => {
    it('should provide project selectors', () => {
      const { result } = renderHook(() => useProjectSelectors());
      
      expect(result.current).toHaveProperty('projects');
      expect(result.current).toHaveProperty('selectedProject');
      expect(result.current).toHaveProperty('isLoading');
      expect(result.current).toHaveProperty('error');
    });

    it('should provide project actions', () => {
      const { result } = renderHook(() => useProjectActions());
      
      expect(result.current).toHaveProperty('setProjects');
      expect(result.current).toHaveProperty('addProject');
      expect(result.current).toHaveProperty('updateProject');
      expect(result.current).toHaveProperty('removeProject');
    });
  });

  describe('Summary Calculation', () => {
    it('should update project summary', () => {
      const { result } = renderHook(() => useProjectStore());
      
      act(() => {
        result.current.setProjects(mockProjects);
        result.current.updateSummary();
      });
      
      const summary = result.current.projectSummary;
      expect(summary).toBeDefined();
      expect(summary!.totalProjects).toBe(2);
      expect(summary!.activeProjects).toBe(1);
      expect(summary!.completedProjects).toBe(1);
      expect(summary!.averageProgress).toBe(75); // (50 + 100) / 2
    });
  });
});
