/**
 * TanStack Query hooks for project data management
 */

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Logger, LogLevel } from '@pnp/logging';

import { 
  getProjects, 
  getProject, 
  createProject, 
  updateProject, 
  deleteProject,
  getProjectSummary,
  bulkUpdateProjects
} from '../services/projectService';
import { 
  Project, 
  ProjectCreateRequest, 
  ProjectUpdateRequest, 
  ProjectFilter,
  ProjectSummary 
} from '../interfaces/Project';
import { useProjectActions } from '../stores/projectStore';
import { getUserFriendlyErrorMessage } from '../helper/errorHelper';

// Query keys
export const projectKeys = {
  all: ['projects'] as const,
  lists: () => [...projectKeys.all, 'list'] as const,
  list: (filters?: ProjectFilter) => [...projectKeys.lists(), filters] as const,
  details: () => [...projectKeys.all, 'detail'] as const,
  detail: (id: number) => [...projectKeys.details(), id] as const,
  summary: () => [...projectKeys.all, 'summary'] as const,
};

/**
 * Hook to fetch all projects with optional filtering
 */
export const useProjects = (filters?: ProjectFilter) => {
  const { setProjects, setLoading, setError, updateSummary } = useProjectActions();

  return useQuery({
    queryKey: projectKeys.list(filters),
    queryFn: () => getProjects(filters),
    staleTime: 5 * 60 * 1000, // 5 minutes
    cacheTime: 10 * 60 * 1000, // 10 minutes
    refetchOnWindowFocus: false,
    retry: (failureCount, error) => {
      // Don't retry on permission errors
      const errorMessage = getUserFriendlyErrorMessage(error);
      if (errorMessage.includes('permission') || errorMessage.includes('access')) {
        return false;
      }
      return failureCount < 2;
    },
    onSuccess: (data) => {
      Logger.write(`Successfully fetched ${data.length} projects`, LogLevel.Info);
      setProjects(data);
      setError(null);
      updateSummary();
    },
    onError: (error) => {
      const errorMessage = getUserFriendlyErrorMessage(error);
      Logger.write(`Error fetching projects: ${errorMessage}`, LogLevel.Error);
      setError(errorMessage);
    },
    onSettled: () => {
      setLoading(false);
    },
  });
};

/**
 * Hook to fetch a single project by ID
 */
export const useProject = (id: number, enabled = true) => {
  return useQuery({
    queryKey: projectKeys.detail(id),
    queryFn: () => getProject(id),
    enabled: enabled && id > 0,
    staleTime: 2 * 60 * 1000, // 2 minutes
    retry: 1,
    onError: (error) => {
      const errorMessage = getUserFriendlyErrorMessage(error);
      Logger.write(`Error fetching project ${id}: ${errorMessage}`, LogLevel.Error);
    },
  });
};

/**
 * Hook to fetch project summary statistics
 */
export const useProjectSummary = () => {
  const { updateSummary } = useProjectActions();

  return useQuery({
    queryKey: projectKeys.summary(),
    queryFn: getProjectSummary,
    staleTime: 5 * 60 * 1000, // 5 minutes
    refetchOnWindowFocus: false,
    onSuccess: (data) => {
      Logger.write('Successfully fetched project summary', LogLevel.Info);
      updateSummary();
    },
    onError: (error) => {
      const errorMessage = getUserFriendlyErrorMessage(error);
      Logger.write(`Error fetching project summary: ${errorMessage}`, LogLevel.Error);
    },
  });
};

/**
 * Hook to create a new project
 */
export const useCreateProject = () => {
  const queryClient = useQueryClient();
  const { addProject, setError } = useProjectActions();

  return useMutation({
    mutationFn: (project: ProjectCreateRequest) => createProject(project),
    onSuccess: (newProject) => {
      Logger.write(`Successfully created project: ${newProject.Name}`, LogLevel.Info);
      
      // Update the store
      addProject(newProject);
      
      // Invalidate and refetch projects
      queryClient.invalidateQueries({ queryKey: projectKeys.lists() });
      queryClient.invalidateQueries({ queryKey: projectKeys.summary() });
      
      setError(null);
    },
    onError: (error) => {
      const errorMessage = getUserFriendlyErrorMessage(error);
      Logger.write(`Error creating project: ${errorMessage}`, LogLevel.Error);
      setError(errorMessage);
    },
  });
};

/**
 * Hook to update an existing project
 */
export const useUpdateProject = () => {
  const queryClient = useQueryClient();
  const { updateProject: updateProjectInStore, setError } = useProjectActions();

  return useMutation({
    mutationFn: (project: ProjectUpdateRequest) => updateProject(project),
    onSuccess: (updatedProject) => {
      Logger.write(`Successfully updated project: ${updatedProject.Name}`, LogLevel.Info);
      
      // Update the store
      updateProjectInStore(updatedProject);
      
      // Update the cache for the specific project
      queryClient.setQueryData(
        projectKeys.detail(updatedProject.Id),
        updatedProject
      );
      
      // Invalidate related queries
      queryClient.invalidateQueries({ queryKey: projectKeys.lists() });
      queryClient.invalidateQueries({ queryKey: projectKeys.summary() });
      
      setError(null);
    },
    onError: (error) => {
      const errorMessage = getUserFriendlyErrorMessage(error);
      Logger.write(`Error updating project: ${errorMessage}`, LogLevel.Error);
      setError(errorMessage);
    },
  });
};

/**
 * Hook to delete a project
 */
export const useDeleteProject = () => {
  const queryClient = useQueryClient();
  const { removeProject, setError } = useProjectActions();

  return useMutation({
    mutationFn: (id: number) => deleteProject(id),
    onSuccess: (_, deletedId) => {
      Logger.write(`Successfully deleted project with ID: ${deletedId}`, LogLevel.Info);
      
      // Update the store
      removeProject(deletedId);
      
      // Remove from cache
      queryClient.removeQueries({ queryKey: projectKeys.detail(deletedId) });
      
      // Invalidate related queries
      queryClient.invalidateQueries({ queryKey: projectKeys.lists() });
      queryClient.invalidateQueries({ queryKey: projectKeys.summary() });
      
      setError(null);
    },
    onError: (error) => {
      const errorMessage = getUserFriendlyErrorMessage(error);
      Logger.write(`Error deleting project: ${errorMessage}`, LogLevel.Error);
      setError(errorMessage);
    },
  });
};

/**
 * Hook for bulk updating projects
 */
export const useBulkUpdateProjects = () => {
  const queryClient = useQueryClient();
  const { setError } = useProjectActions();

  return useMutation({
    mutationFn: (updates: ProjectUpdateRequest[]) => bulkUpdateProjects(updates),
    onSuccess: (updatedProjects) => {
      Logger.write(`Successfully bulk updated ${updatedProjects.length} projects`, LogLevel.Info);
      
      // Update individual project caches
      updatedProjects.forEach(project => {
        queryClient.setQueryData(
          projectKeys.detail(project.Id),
          project
        );
      });
      
      // Invalidate list queries to trigger refetch
      queryClient.invalidateQueries({ queryKey: projectKeys.lists() });
      queryClient.invalidateQueries({ queryKey: projectKeys.summary() });
      
      setError(null);
    },
    onError: (error) => {
      const errorMessage = getUserFriendlyErrorMessage(error);
      Logger.write(`Error bulk updating projects: ${errorMessage}`, LogLevel.Error);
      setError(errorMessage);
    },
  });
};

/**
 * Hook to prefetch a project (useful for hover states)
 */
export const usePrefetchProject = () => {
  const queryClient = useQueryClient();

  return (id: number) => {
    queryClient.prefetchQuery({
      queryKey: projectKeys.detail(id),
      queryFn: () => getProject(id),
      staleTime: 2 * 60 * 1000, // 2 minutes
    });
  };
};

/**
 * Hook to get cached project data without triggering a fetch
 */
export const useCachedProject = (id: number) => {
  const queryClient = useQueryClient();
  
  return queryClient.getQueryData<Project>(projectKeys.detail(id));
};

/**
 * Hook to manually invalidate project queries
 */
export const useInvalidateProjects = () => {
  const queryClient = useQueryClient();

  return {
    invalidateAll: () => queryClient.invalidateQueries({ queryKey: projectKeys.all }),
    invalidateLists: () => queryClient.invalidateQueries({ queryKey: projectKeys.lists() }),
    invalidateProject: (id: number) => queryClient.invalidateQueries({ queryKey: projectKeys.detail(id) }),
    invalidateSummary: () => queryClient.invalidateQueries({ queryKey: projectKeys.summary() }),
  };
};
