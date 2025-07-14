/**
 * TanStack Query hooks for project data management - SharePoint 2019 On-Premise Compatible
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