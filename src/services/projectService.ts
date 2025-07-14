/**
 * Enhanced project service with full CRUD operations
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import { Logger, LogLevel } from "@pnp/logging";

import { Project, ProjectCreateRequest, ProjectUpdateRequest, ProjectFilter, ProjectSummary } from '../interfaces/Project';
import { Status } from '../interfaces/Status';
import { getErrorMessage, logError, handleAsyncOperation } from '../helper/errorHelper';

const LIST_TITLE = "Projects";

/**
 * Maps SharePoint list item to Project interface
 */
const mapItemToProject = (item: any): Project => ({
  Id: item.Id,
  Name: item.Title,
  Status: item.Status || Status.NotStarted,
  Description: item.Description,
  StartDate: item.StartDate ? new Date(item.StartDate) : undefined,
  EndDate: item.EndDate ? new Date(item.EndDate) : undefined,
  Owner: item.Owner,
  Priority: item.Priority || 0,
  Progress: item.Progress || 0,
  Budget: item.Budget,
  Tags: item.Tags ? item.Tags.split(';') : [],
  CreatedDate: new Date(item.Created),
  ModifiedDate: new Date(item.Modified),
  CreatedBy: item.Author?.Title,
  ModifiedBy: item.Editor?.Title,
});

/**
 * Gets all projects with optional filtering
 */
export const getProjects = async (filter?: ProjectFilter): Promise<Project[]> => {
  return handleAsyncOperation(async () => {
    let query = sp.web.lists.getByTitle(LIST_TITLE).items
      .select(
        "Id", "Title", "Status", "Description", "StartDate", "EndDate", 
        "Owner", "Priority", "Progress", "Budget", "Tags",
        "Created", "Modified", "Author/Title", "Editor/Title"
      )
      .expand("Author", "Editor")
      .top(100);

    // Apply filters
    if (filter) {
      const filterConditions: string[] = [];

      if (filter.status !== undefined) {
        filterConditions.push(`Status eq ${filter.status}`);
      }

      if (filter.priority !== undefined) {
        filterConditions.push(`Priority eq ${filter.priority}`);
      }

      if (filter.owner) {
        filterConditions.push(`Owner eq '${filter.owner}'`);
      }

      if (filter.searchText) {
        filterConditions.push(`substringof('${filter.searchText}', Title)`);
      }

      if (filter.startDateFrom) {
        filterConditions.push(`StartDate ge datetime'${filter.startDateFrom.toISOString()}'`);
      }

      if (filter.startDateTo) {
        filterConditions.push(`StartDate le datetime'${filter.startDateTo.toISOString()}'`);
      }

      if (filterConditions.length > 0) {
        query = query.filter(filterConditions.join(' and '));
      }
    }

    const items = await query.get();
    return items.map(mapItemToProject);
  }, 'getProjects');
};

/**
 * Gets a single project by ID
 */
export const getProject = async (id: number): Promise<Project | null> => {
  return handleAsyncOperation(async () => {
    const item = await sp.web.lists.getByTitle(LIST_TITLE).items
      .getById(id)
      .select(
        "Id", "Title", "Status", "Description", "StartDate", "EndDate", 
        "Owner", "Priority", "Progress", "Budget", "Tags",
        "Created", "Modified", "Author/Title", "Editor/Title"
      )
      .expand("Author", "Editor")
      .get();

    return mapItemToProject(item);
  }, `getProject(${id})`);
};

/**
 * Creates a new project
 */
export const createProject = async (project: ProjectCreateRequest): Promise<Project> => {
  return handleAsyncOperation(async () => {
    const itemData: any = {
      Title: project.Name,
      Status: project.Status || Status.NotStarted,
      Description: project.Description,
      StartDate: project.StartDate?.toISOString(),
      EndDate: project.EndDate?.toISOString(),
      Owner: project.Owner,
      Priority: project.Priority || 0,
      Budget: project.Budget,
      Tags: project.Tags?.join(';'),
    };

    const result = await sp.web.lists.getByTitle(LIST_TITLE).items.add(itemData);
    
    // Get the created item with all fields
    const createdItem = await sp.web.lists.getByTitle(LIST_TITLE).items
      .getById(result.data.Id)
      .select(
        "Id", "Title", "Status", "Description", "StartDate", "EndDate", 
        "Owner", "Priority", "Progress", "Budget", "Tags",
        "Created", "Modified", "Author/Title", "Editor/Title"
      )
      .expand("Author", "Editor")
      .get();

    return mapItemToProject(createdItem);
  }, 'createProject');
};

/**
 * Updates an existing project
 */
export const updateProject = async (project: ProjectUpdateRequest): Promise<Project> => {
  return handleAsyncOperation(async () => {
    const itemData: any = {};

    if (project.Name !== undefined) itemData.Title = project.Name;
    if (project.Status !== undefined) itemData.Status = project.Status;
    if (project.Description !== undefined) itemData.Description = project.Description;
    if (project.StartDate !== undefined) itemData.StartDate = project.StartDate?.toISOString();
    if (project.EndDate !== undefined) itemData.EndDate = project.EndDate?.toISOString();
    if (project.Owner !== undefined) itemData.Owner = project.Owner;
    if (project.Priority !== undefined) itemData.Priority = project.Priority;
    if (project.Progress !== undefined) itemData.Progress = project.Progress;
    if (project.Budget !== undefined) itemData.Budget = project.Budget;
    if (project.Tags !== undefined) itemData.Tags = project.Tags?.join(';');

    await sp.web.lists.getByTitle(LIST_TITLE).items.getById(project.Id).update(itemData);
    
    // Get the updated item
    const updatedItem = await sp.web.lists.getByTitle(LIST_TITLE).items
      .getById(project.Id)
      .select(
        "Id", "Title", "Status", "Description", "StartDate", "EndDate", 
        "Owner", "Priority", "Progress", "Budget", "Tags",
        "Created", "Modified", "Author/Title", "Editor/Title"
      )
      .expand("Author", "Editor")
      .get();

    return mapItemToProject(updatedItem);
  }, `updateProject(${project.Id})`);
};

/**
 * Deletes a project
 */
export const deleteProject = async (id: number): Promise<void> => {
  return handleAsyncOperation(async () => {
    await sp.web.lists.getByTitle(LIST_TITLE).items.getById(id).delete();
  }, `deleteProject(${id})`);
};

/**
 * Gets project summary statistics
 */
export const getProjectSummary = async (): Promise<ProjectSummary> => {
  return handleAsyncOperation(async () => {
    const projects = await getProjects();
    
    const summary: ProjectSummary = {
      totalProjects: projects.length,
      activeProjects: projects.filter(p => p.Status === Status.InProgress).length,
      completedProjects: projects.filter(p => p.Status === Status.Completed).length,
      onHoldProjects: projects.filter(p => p.Status === Status.OnHold).length,
      notStartedProjects: projects.filter(p => p.Status === Status.NotStarted).length,
      averageProgress: projects.length > 0 
        ? projects.reduce((sum, p) => sum + (p.Progress || 0), 0) / projects.length 
        : 0,
      totalBudget: projects.reduce((sum, p) => sum + (p.Budget || 0), 0),
    };

    return summary;
  }, 'getProjectSummary');
};

/**
 * Bulk updates multiple projects
 */
export const bulkUpdateProjects = async (updates: ProjectUpdateRequest[]): Promise<Project[]> => {
  return handleAsyncOperation(async () => {
    const batch = sp.web.createBatch();
    const list = sp.web.lists.getByTitle(LIST_TITLE);

    // Add all updates to the batch
    updates.forEach(update => {
      const itemData: any = {};
      
      if (update.Name !== undefined) itemData.Title = update.Name;
      if (update.Status !== undefined) itemData.Status = update.Status;
      if (update.Description !== undefined) itemData.Description = update.Description;
      if (update.Progress !== undefined) itemData.Progress = update.Progress;
      
      list.items.getById(update.Id).inBatch(batch).update(itemData);
    });

    await batch.execute();

    // Get updated items
    const updatedProjects: Project[] = [];
    for (const update of updates) {
      const project = await getProject(update.Id);
      if (project) {
        updatedProjects.push(project);
      }
    }

    return updatedProjects;
  }, 'bulkUpdateProjects');
};
