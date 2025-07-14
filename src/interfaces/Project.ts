/**
 * Enhanced Project interface with additional properties
 */

import { Status } from './Status';

export interface Project {
  Id: number;
  Name: string;
  Status: Status;
  Description?: string;
  StartDate?: Date;
  EndDate?: Date;
  Owner?: string;
  Priority?: ProjectPriority;
  Progress?: number; // 0-100
  Budget?: number;
  Tags?: string[];
  CreatedDate?: Date;
  ModifiedDate?: Date;
  CreatedBy?: string;
  ModifiedBy?: string;
}

export enum ProjectPriority {
  Low = 0,
  Medium = 1,
  High = 2,
  Critical = 3,
}

export interface ProjectSummary {
  totalProjects: number;
  activeProjects: number;
  completedProjects: number;
  onHoldProjects: number;
  notStartedProjects: number;
  averageProgress: number;
  totalBudget: number;
}

export interface ProjectFilter {
  status?: Status;
  priority?: ProjectPriority;
  owner?: string;
  startDateFrom?: Date;
  startDateTo?: Date;
  searchText?: string;
  tags?: string[];
}

export interface ProjectCreateRequest {
  Name: string;
  Description?: string;
  Status?: Status;
  StartDate?: Date;
  EndDate?: Date;
  Owner?: string;
  Priority?: ProjectPriority;
  Budget?: number;
  Tags?: string[];
}

export interface ProjectUpdateRequest extends Partial<ProjectCreateRequest> {
  Id: number;
  Progress?: number;
}
