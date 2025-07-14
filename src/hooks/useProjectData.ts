/**
 * Custom hook for project data management - React 17 with SPFx 1.4.1
 * Modern React 17 patterns with SharePoint 2019 compatibility
 */

import { useState, useEffect, useCallback } from 'react';
import { Logger, LogLevel } from '@pnp/logging';
import { useProjectStore } from '../stores/projectStore';
import { 
  getProjects, 
  getProject, 
  createProject, 
  updateProject, 
  deleteProject,
  getProjectSummary,
  bulkUpdateProjects
} from '../services/projectService';
import { Project, ProjectFilter } from '../interfaces/Project';

// Custom query state interface
interface QueryState<T> {
  data: T | null;
  isLoading: boolean;
  error: string | null;
  isError: boolean;
  isSuccess: boolean;
}

// Custom mutation state interface
interface MutationState<T> {
  data: T | null;
  isLoading: boolean;
  error: string | null;
  isError: boolean;
  isSuccess: boolean;
  mutate: (variables: any) => Promise<void>;
  reset: () => void;
}

// Custom hook for projects query with React 17 patterns
export function useProjectsQuery(filters?: ProjectFilter): QueryState<Project[]> {
  const [state, setState] = useState<QueryState<Project[]>>({
    data: null,
    isLoading: true,
    error: null,
    isError: false,
    isSuccess: false,
  });

  const { setProjects, setError, setLoading } = useProjectStore();

  useEffect(() => {
    let isCancelled = false;

    const fetchProjects = async () => {
      try {
        setState(prev => ({ ...prev, isLoading: true, error: null, isError: false }));
        setLoading(true);
        
        const projects = await getProjects(filters);
        
        if (!isCancelled) {
          setState({
            data: projects,
            isLoading: false,
            error: null,
            isError: false,
            isSuccess: true,
          });
          
          setProjects(projects);
          setLoading(false);
        }
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Failed to fetch projects';
        Logger.write(`Error fetching projects: ${errorMessage}`, LogLevel.Error);
        
        if (!isCancelled) {
          setState({
            data: null,
            isLoading: false,
            error: errorMessage,
            isError: true,
            isSuccess: false,
          });
          
          setError(errorMessage);
        }
      }
    };

    fetchProjects();

    return () => {
      isCancelled = true;
    };
  }, [filters ? JSON.stringify(filters) : '', setProjects, setError, setLoading]);

  return state;
}

// Custom hook for single project query with React 17
export function useProjectQuery(projectId: number): QueryState<Project> {
  const [state, setState] = useState<QueryState<Project>>({
    data: null,
    isLoading: true,
    error: null,
    isError: false,
    isSuccess: false,
  });

  useEffect(() => {
    if (!projectId) {
      setState({
        data: null,
        isLoading: false,
        error: null,
        isError: false,
        isSuccess: true,
      });
      return;
    }

    let isCancelled = false;

    const fetchProject = async () => {
      try {
        setState(prev => ({ ...prev, isLoading: true, error: null, isError: false }));

        const project = await getProject(projectId);
        
        if (!isCancelled) {
          setState({
            data: project,
            isLoading: false,
            error: null,
            isError: false,
            isSuccess: true,
          });
        }
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Failed to fetch project';
        Logger.write(`Error fetching project: ${errorMessage}`, LogLevel.Error);
        
        if (!isCancelled) {
          setState({
            data: null,
            isLoading: false,
            error: errorMessage,
            isError: true,
            isSuccess: false,
          });
        }
      }
    };

    fetchProject();

    return () => {
      isCancelled = true;
    };
  }, [projectId]);

  return state;
}

// Custom hook for create project mutation with React 17
export function useCreateProjectMutation(): MutationState<Project> {
  const [state, setState] = useState<MutationState<Project>>({
    data: null,
    isLoading: false,
    error: null,
    isError: false,
    isSuccess: false,
    mutate: async () => {},
    reset: () => {},
  });

  const { addProject } = useProjectStore();

  const mutate = useCallback(async (projectData: Partial<Project>) => {
    try {
      setState(prev => ({ 
        ...prev, 
        isLoading: true, 
        error: null, 
        isError: false, 
        isSuccess: false 
      }));

      const newProject = await createProject(projectData);
      
      setState(prev => ({
        ...prev,
        data: newProject,
        isLoading: false,
        error: null,
        isError: false,
        isSuccess: true,
      }));
      
      addProject(newProject);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to create project';
      Logger.write(`Error creating project: ${errorMessage}`, LogLevel.Error);
      
      setState(prev => ({
        ...prev,
        isLoading: false,
        error: errorMessage,
        isError: true,
        isSuccess: false,
      }));
      
      throw error;
    }
  }, [addProject]);

  const reset = useCallback(() => {
    setState(prev => ({
      ...prev,
      data: null,
      error: null,
      isError: false,
      isSuccess: false,
    }));
  }, []);

  useEffect(() => {
    setState(prev => ({
      ...prev,
      mutate,
      reset,
    }));
  }, [mutate, reset]);

  return state;
}

// Custom hook for update project mutation with React 17
export function useUpdateProjectMutation(): MutationState<Project> {
  const [state, setState] = useState<MutationState<Project>>({
    data: null,
    isLoading: false,
    error: null,
    isError: false,
    isSuccess: false,
    mutate: async () => {},
    reset: () => {},
  });

  const { updateProject: updateProjectInStore } = useProjectStore();

  const mutate = useCallback(async (projectData: Project) => {
    try {
      setState(prev => ({ 
        ...prev, 
        isLoading: true, 
        error: null, 
        isError: false, 
        isSuccess: false 
      }));

      const updatedProject = await updateProject(projectData);
      
      setState(prev => ({
        ...prev,
        data: updatedProject,
        isLoading: false,
        error: null,
        isError: false,
        isSuccess: true,
      }));
      
      updateProjectInStore(updatedProject);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to update project';
      Logger.write(`Error updating project: ${errorMessage}`, LogLevel.Error);
      
      setState(prev => ({
        ...prev,
        isLoading: false,
        error: errorMessage,
        isError: true,
        isSuccess: false,
      }));
      
      throw error;
    }
  }, [updateProjectInStore]);

  const reset = useCallback(() => {
    setState(prev => ({
      ...prev,
      data: null,
      error: null,
      isError: false,
      isSuccess: false,
    }));
  }, []);

  useEffect(() => {
    setState(prev => ({
      ...prev,
      mutate,
      reset,
    }));
  }, [mutate, reset]);

  return state;
}

// Custom hook for delete project mutation with React 17
export function useDeleteProjectMutation(): MutationState<void> {
  const [state, setState] = useState<MutationState<void>>({
    data: null,
    isLoading: false,
    error: null,
    isError: false,
    isSuccess: false,
    mutate: async () => {},
    reset: () => {},
  });

  const { removeProject } = useProjectStore();

  const mutate = useCallback(async (projectId: number) => {
    try {
      setState(prev => ({ 
        ...prev, 
        isLoading: true, 
        error: null, 
        isError: false, 
        isSuccess: false 
      }));

      await deleteProject(projectId);
      
      setState(prev => ({
        ...prev,
        data: null,
        isLoading: false,
        error: null,
        isError: false,
        isSuccess: true,
      }));
      
      removeProject(projectId);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to delete project';
      Logger.write(`Error deleting project: ${errorMessage}`, LogLevel.Error);
      
      setState(prev => ({
        ...prev,
        isLoading: false,
        error: errorMessage,
        isError: true,
        isSuccess: false,
      }));
      
      throw error;
    }
  }, [removeProject]);

  const reset = useCallback(() => {
    setState(prev => ({
      ...prev,
      data: null,
      error: null,
      isError: false,
      isSuccess: false,
    }));
  }, []);

  useEffect(() => {
    setState(prev => ({
      ...prev,
      mutate,
      reset,
    }));
  }, [mutate, reset]);

  return state;
}

// Custom hook for project summary with React 17
export function useProjectSummaryQuery(): QueryState<any> {
  const [state, setState] = useState<QueryState<any>>({
    data: null,
    isLoading: true,
    error: null,
    isError: false,
    isSuccess: false,
  });

  useEffect(() => {
    let isCancelled = false;

    const fetchSummary = async () => {
      try {
        setState(prev => ({ ...prev, isLoading: true, error: null, isError: false }));

        const summary = await getProjectSummary();
        
        if (!isCancelled) {
          setState({
            data: summary,
            isLoading: false,
            error: null,
            isError: false,
            isSuccess: true,
          });
        }
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Failed to fetch project summary';
        Logger.write(`Error fetching project summary: ${errorMessage}`, LogLevel.Error);
        
        if (!isCancelled) {
          setState({
            data: null,
            isLoading: false,
            error: errorMessage,
            isError: true,
            isSuccess: false,
          });
        }
      }
    };

    fetchSummary();

    return () => {
      isCancelled = true;
    };
  }, []);

  return state;
}

// Bulk operations hook with React 17
export function useBulkUpdateMutation(): MutationState<Project[]> {
  const [state, setState] = useState<MutationState<Project[]>>({
    data: null,
    isLoading: false,
    error: null,
    isError: false,
    isSuccess: false,
    mutate: async () => {},
    reset: () => {},
  });

  const { setProjects } = useProjectStore();

  const mutate = useCallback(async (updates: { projectIds: number[]; updates: Partial<Project> }) => {
    try {
      setState(prev => ({ 
        ...prev, 
        isLoading: true, 
        error: null, 
        isError: false, 
        isSuccess: false 
      }));

      const updatedProjects = await bulkUpdateProjects(updates.projectIds, updates.updates);
      
      setState(prev => ({
        ...prev,
        data: updatedProjects,
        isLoading: false,
        error: null,
        isError: false,
        isSuccess: true,
      }));
      
      // Refresh projects list
      const allProjects = await getProjects();
      setProjects(allProjects);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to bulk update projects';
      Logger.write(`Error bulk updating projects: ${errorMessage}`, LogLevel.Error);
      
      setState(prev => ({
        ...prev,
        isLoading: false,
        error: errorMessage,
        isError: true,
        isSuccess: false,
      }));
      
      throw error;
    }
  }, [setProjects]);

  const reset = useCallback(() => {
    setState(prev => ({
      ...prev,
      data: null,
      error: null,
      isError: false,
      isSuccess: false,
    }));
  }, []);

  useEffect(() => {
    setState(prev => ({
      ...prev,
      mutate,
      reset,
    }));
  }, [mutate, reset]);

  return state;
}
