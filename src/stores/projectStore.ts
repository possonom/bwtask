/**
 * Zustand store for project management - React 17 with SPFx 1.4.1
 */

import { create } from 'zustand';
import { subscribeWithSelector } from 'zustand/middleware';
import { Project, ProjectFilter, ProjectSummary, ProjectPriority } from '../interfaces/Project';
import { Status } from '../interfaces/Status';

export interface ProjectState {
  // Data
  projects: Project[];
  selectedProject: Project | null;
  projectSummary: ProjectSummary | null;
  
  // UI State
  isLoading: boolean;
  error: string | null;
  filters: ProjectFilter;
  searchText: string;
  sortBy: 'name' | 'status' | 'priority' | 'startDate' | 'progress';
  sortDirection: 'asc' | 'desc';
  
  // Panel/Modal state
  isPanelOpen: boolean;
  isCreateModalOpen: boolean;
  isEditMode: boolean;
  
  // Actions
  setProjects: (projects: Project[]) => void;
  addProject: (project: Project) => void;
  updateProject: (project: Project) => void;
  removeProject: (projectId: number) => void;
  setSelectedProject: (project: Project | null) => void;
  
  // UI Actions
  setLoading: (loading: boolean) => void;
  setError: (error: string | null) => void;
  setFilters: (filters: Partial<ProjectFilter>) => void;
  setSearchText: (text: string) => void;
  setSorting: (sortBy: ProjectState['sortBy'], direction: ProjectState['sortDirection']) => void;
  
  // Panel/Modal actions
  openPanel: (project?: Project) => void;
  closePanel: () => void;
  openCreateModal: () => void;
  closeCreateModal: () => void;
  setEditMode: (editMode: boolean) => void;
  
  // Computed values
  getFilteredProjects: () => Project[];
  getSortedProjects: () => Project[];
  
  // Summary actions
  updateSummary: () => void;
  
  // Bulk actions
  bulkUpdateStatus: (projectIds: number[], status: Status) => void;
  bulkUpdatePriority: (projectIds: number[], priority: ProjectPriority) => void;
}

const initialFilters: ProjectFilter = {
  status: undefined,
  priority: undefined,
  owner: undefined,
  searchText: undefined,
  startDateFrom: undefined,
  startDateTo: undefined,
  tags: undefined,
};

// Helper functions for computed values (React 17 compatible)
const filterProjects = (projects: Project[], filters: ProjectFilter, searchText: string): Project[] => {
  return projects.filter((project) => {
    // Search text filter
    if (searchText && !project.Name.toLowerCase().includes(searchText.toLowerCase())) {
      return false;
    }
    
    // Status filter
    if (filters.status !== undefined && project.Status !== filters.status) {
      return false;
    }
    
    // Priority filter
    if (filters.priority !== undefined && project.Priority !== filters.priority) {
      return false;
    }
    
    // Owner filter
    if (filters.owner && project.Owner !== filters.owner) {
      return false;
    }
    
    // Date filters
    if (filters.startDateFrom && project.StartDate && project.StartDate < filters.startDateFrom) {
      return false;
    }
    
    if (filters.startDateTo && project.StartDate && project.StartDate > filters.startDateTo) {
      return false;
    }
    
    // Tags filter
    if (filters.tags && filters.tags.length > 0) {
      const projectTags = project.Tags || [];
      const hasMatchingTag = filters.tags.some(filterTag =>
        projectTags.some(projectTag =>
          projectTag.toLowerCase().includes(filterTag.toLowerCase())
        )
      );
      
      if (!hasMatchingTag) {
        return false;
      }
    }
    
    return true;
  });
};

const sortProjects = (projects: Project[], sortBy: string, sortDirection: string): Project[] => {
  return [...projects].sort((a, b) => {
    let aValue: any;
    let bValue: any;
    
    switch (sortBy) {
      case 'name':
        aValue = a.Name.toLowerCase();
        bValue = b.Name.toLowerCase();
        break;
      case 'status':
        aValue = a.Status;
        bValue = b.Status;
        break;
      case 'priority':
        aValue = a.Priority || 0;
        bValue = b.Priority || 0;
        break;
      case 'startDate':
        aValue = a.StartDate ? a.StartDate.getTime() : 0;
        bValue = b.StartDate ? b.StartDate.getTime() : 0;
        break;
      case 'progress':
        aValue = a.Progress || 0;
        bValue = b.Progress || 0;
        break;
      default:
        return 0;
    }
    
    if (aValue < bValue) return sortDirection === 'asc' ? -1 : 1;
    if (aValue > bValue) return sortDirection === 'asc' ? 1 : -1;
    return 0;
  });
};

export const useProjectStore = create<ProjectState>()(
  subscribeWithSelector((set, get) => ({
    // Initial state
    projects: [],
    selectedProject: null,
    projectSummary: null,
    isLoading: false,
    error: null,
    filters: initialFilters,
    searchText: '',
    sortBy: 'name',
    sortDirection: 'asc',
    isPanelOpen: false,
    isCreateModalOpen: false,
    isEditMode: false,

    // Data actions
    setProjects: (projects) => {
      set((state) => ({
        ...state,
        projects,
        error: null,
      }));
    },

    addProject: (project) => {
      set((state) => ({
        ...state,
        projects: [...state.projects, project],
      }));
    },

    updateProject: (updatedProject) => {
      set((state) => ({
        ...state,
        projects: state.projects.map((p) =>
          p.Id === updatedProject.Id ? updatedProject : p
        ),
        selectedProject: state.selectedProject?.Id === updatedProject.Id 
          ? updatedProject 
          : state.selectedProject,
      }));
    },

    removeProject: (projectId) => {
      set((state) => ({
        ...state,
        projects: state.projects.filter((p) => p.Id !== projectId),
        selectedProject: state.selectedProject?.Id === projectId 
          ? null 
          : state.selectedProject,
      }));
    },

    setSelectedProject: (project) => {
      set((state) => ({
        ...state,
        selectedProject: project,
      }));
    },

    // UI actions
    setLoading: (loading) => {
      set((state) => ({
        ...state,
        isLoading: loading,
      }));
    },

    setError: (error) => {
      set((state) => ({
        ...state,
        error,
        isLoading: false,
      }));
    },

    setFilters: (newFilters) => {
      set((state) => ({
        ...state,
        filters: { ...state.filters, ...newFilters },
      }));
    },

    setSearchText: (text) => {
      set((state) => ({
        ...state,
        searchText: text,
        filters: { ...state.filters, searchText: text },
      }));
    },

    setSorting: (sortBy, direction) => {
      set((state) => ({
        ...state,
        sortBy,
        sortDirection: direction,
      }));
    },

    // Panel/Modal actions
    openPanel: (project) => {
      set((state) => ({
        ...state,
        isPanelOpen: true,
        selectedProject: project || null,
        isEditMode: !!project,
      }));
    },

    closePanel: () => {
      set((state) => ({
        ...state,
        isPanelOpen: false,
        selectedProject: null,
        isEditMode: false,
      }));
    },

    openCreateModal: () => {
      set((state) => ({
        ...state,
        isCreateModalOpen: true,
      }));
    },

    closeCreateModal: () => {
      set((state) => ({
        ...state,
        isCreateModalOpen: false,
      }));
    },

    setEditMode: (editMode) => {
      set((state) => ({
        ...state,
        isEditMode: editMode,
      }));
    },

    // Computed values as functions for React 17
    getFilteredProjects: () => {
      const state = get();
      return filterProjects(state.projects, state.filters, state.searchText);
    },

    getSortedProjects: () => {
      const state = get();
      const filtered = filterProjects(state.projects, state.filters, state.searchText);
      return sortProjects(filtered, state.sortBy, state.sortDirection);
    },

    // Summary actions
    updateSummary: () => {
      set((state) => {
        const { projects } = state;
        
        const totalProjects = projects.length;
        const activeProjects = projects.filter(p => p.Status === Status.InProgress).length;
        const completedProjects = projects.filter(p => p.Status === Status.Completed).length;
        const onHoldProjects = projects.filter(p => p.Status === Status.OnHold).length;
        const notStartedProjects = projects.filter(p => p.Status === Status.NotStarted).length;
        
        const totalProgress = projects.reduce((sum, p) => sum + (p.Progress || 0), 0);
        const totalBudget = projects.reduce((sum, p) => sum + (p.Budget || 0), 0);
        
        const summary: ProjectSummary = {
          totalProjects,
          activeProjects,
          completedProjects,
          onHoldProjects,
          notStartedProjects,
          averageProgress: totalProjects > 0 ? totalProgress / totalProjects : 0,
          totalBudget,
        };

        return {
          ...state,
          projectSummary: summary,
        };
      });
    },

    // Bulk actions
    bulkUpdateStatus: (projectIds, status) => {
      set((state) => ({
        ...state,
        projects: state.projects.map((project) =>
          projectIds.includes(project.Id)
            ? { ...project, Status: status, ModifiedDate: new Date() }
            : project
        ),
      }));
    },

    bulkUpdatePriority: (projectIds, priority) => {
      set((state) => ({
        ...state,
        projects: state.projects.map((project) =>
          projectIds.includes(project.Id)
            ? { ...project, Priority: priority, ModifiedDate: new Date() }
            : project
        ),
      }));
    },
  }))
);
