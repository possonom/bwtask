/**
 * Project service for SharePoint data operations - SPFx 1.4.1 Compatible
 */

import { sp } from '@pnp/sp';
import { Logger, LogLevel } from '@pnp/logging';
import { Project, ProjectFilter, ProjectSummary } from '../interfaces/Project';
import { Status } from '../interfaces/Status';

// Initialize PnP SP for SharePoint 2019
export function initializeProjectService(context: any): void {
  try {
    sp.setup({
      spfxContext: context,
      sp: {
        headers: {
          'Accept': 'application/json; odata=verbose',
          'Content-Type': 'application/json; odata=verbose',
        },
        baseUrl: context.pageContext.web.absoluteUrl
      }
    });
    
    Logger.write('Project service initialized for SharePoint 2019', LogLevel.Info);
  } catch (error) {
    Logger.write('Error initializing project service: ' + (error instanceof Error ? error.message : String(error)), LogLevel.Error);
    throw error;
  }
}

// Get all projects with optional filtering
export async function getProjects(filter?: ProjectFilter): Promise<Project[]> {
  try {
    Logger.write('Fetching projects from SharePoint 2019', LogLevel.Info);
    
    let query = sp.web.lists.getByTitle('Projects').items
      .select(
        'Id',
        'Title',
        'Description',
        'Status',
        'Priority',
        'StartDate',
        'EndDate',
        'Progress',
        'Budget',
        'Owner',
        'Tags',
        'Created',
        'Modified',
        'Author/Title',
        'Editor/Title'
      )
      .expand('Author', 'Editor')
      .top(1000); // SharePoint 2019 limit

    // Apply filters if provided
    if (filter) {
      var filterConditions: string[] = [];
      
      if (filter.status !== undefined) {
        filterConditions.push('Status eq \'' + filter.status + '\'');
      }
      
      if (filter.priority !== undefined) {
        filterConditions.push('Priority eq ' + filter.priority);
      }
      
      if (filter.owner) {
        filterConditions.push('Owner eq \'' + filter.owner + '\'');
      }
      
      if (filter.startDateFrom) {
        filterConditions.push('StartDate ge datetime\'' + filter.startDateFrom.toISOString() + '\'');
      }
      
      if (filter.startDateTo) {
        filterConditions.push('StartDate le datetime\'' + filter.startDateTo.toISOString() + '\'');
      }
      
      if (filterConditions.length > 0) {
        query = query.filter(filterConditions.join(' and '));
      }
    }

    const items = await query.get();
    
    const projects: Project[] = items.map(function(item: any): Project {
      return {
        Id: item.Id,
        Name: item.Title || '',
        Description: item.Description || '',
        Status: item.Status || Status.NotStarted,
        Priority: item.Priority || 1,
        StartDate: item.StartDate ? new Date(item.StartDate) : null,
        EndDate: item.EndDate ? new Date(item.EndDate) : null,
        Progress: item.Progress || 0,
        Budget: item.Budget || 0,
        Owner: item.Owner || '',
        Tags: item.Tags ? item.Tags.split(',').map(function(tag: string) { return tag.trim(); }) : [],
        CreatedDate: new Date(item.Created),
        ModifiedDate: new Date(item.Modified),
        CreatedBy: item.Author ? item.Author.Title : '',
        ModifiedBy: item.Editor ? item.Editor.Title : '',
      };
    });

    Logger.write('Successfully fetched ' + projects.length + ' projects', LogLevel.Info);
    return projects;
    
  } catch (error) {
    const errorMessage = 'Error fetching projects: ' + (error instanceof Error ? error.message : String(error));
    Logger.write(errorMessage, LogLevel.Error);
    throw new Error(errorMessage);
  }
}

// Get a single project by ID
export async function getProject(projectId: number): Promise<Project> {
  try {
    Logger.write('Fetching project with ID: ' + projectId, LogLevel.Info);
    
    const item = await sp.web.lists.getByTitle('Projects').items
      .getById(projectId)
      .select(
        'Id',
        'Title',
        'Description',
        'Status',
        'Priority',
        'StartDate',
        'EndDate',
        'Progress',
        'Budget',
        'Owner',
        'Tags',
        'Created',
        'Modified',
        'Author/Title',
        'Editor/Title'
      )
      .expand('Author', 'Editor')
      .get();

    const project: Project = {
      Id: item.Id,
      Name: item.Title || '',
      Description: item.Description || '',
      Status: item.Status || Status.NotStarted,
      Priority: item.Priority || 1,
      StartDate: item.StartDate ? new Date(item.StartDate) : null,
      EndDate: item.EndDate ? new Date(item.EndDate) : null,
      Progress: item.Progress || 0,
      Budget: item.Budget || 0,
      Owner: item.Owner || '',
      Tags: item.Tags ? item.Tags.split(',').map(function(tag: string) { return tag.trim(); }) : [],
      CreatedDate: new Date(item.Created),
      ModifiedDate: new Date(item.Modified),
      CreatedBy: item.Author ? item.Author.Title : '',
      ModifiedBy: item.Editor ? item.Editor.Title : '',
    };

    Logger.write('Successfully fetched project: ' + project.Name, LogLevel.Info);
    return project;
    
  } catch (error) {
    const errorMessage = 'Error fetching project: ' + (error instanceof Error ? error.message : String(error));
    Logger.write(errorMessage, LogLevel.Error);
    throw new Error(errorMessage);
  }
}

// Create a new project
export async function createProject(projectData: Partial<Project>): Promise<Project> {
  try {
    Logger.write('Creating new project: ' + (projectData.Name || 'Unnamed'), LogLevel.Info);
    
    const itemData = {
      Title: projectData.Name || '',
      Description: projectData.Description || '',
      Status: projectData.Status || Status.NotStarted,
      Priority: projectData.Priority || 1,
      StartDate: projectData.StartDate ? projectData.StartDate.toISOString() : null,
      EndDate: projectData.EndDate ? projectData.EndDate.toISOString() : null,
      Progress: projectData.Progress || 0,
      Budget: projectData.Budget || 0,
      Owner: projectData.Owner || '',
      Tags: projectData.Tags ? projectData.Tags.join(', ') : '',
    };

    const result = await sp.web.lists.getByTitle('Projects').items.add(itemData);
    
    // Fetch the created item with all fields
    const createdProject = await getProject(result.data.Id);
    
    Logger.write('Successfully created project with ID: ' + createdProject.Id, LogLevel.Info);
    return createdProject;
    
  } catch (error) {
    const errorMessage = 'Error creating project: ' + (error instanceof Error ? error.message : String(error));
    Logger.write(errorMessage, LogLevel.Error);
    throw new Error(errorMessage);
  }
}

// Update an existing project
export async function updateProject(project: Project): Promise<Project> {
  try {
    Logger.write('Updating project: ' + project.Name + ' (ID: ' + project.Id + ')', LogLevel.Info);
    
    const itemData = {
      Title: project.Name,
      Description: project.Description,
      Status: project.Status,
      Priority: project.Priority,
      StartDate: project.StartDate ? project.StartDate.toISOString() : null,
      EndDate: project.EndDate ? project.EndDate.toISOString() : null,
      Progress: project.Progress,
      Budget: project.Budget,
      Owner: project.Owner,
      Tags: project.Tags ? project.Tags.join(', ') : '',
    };

    await sp.web.lists.getByTitle('Projects').items.getById(project.Id).update(itemData);
    
    // Fetch the updated item
    const updatedProject = await getProject(project.Id);
    
    Logger.write('Successfully updated project: ' + updatedProject.Name, LogLevel.Info);
    return updatedProject;
    
  } catch (error) {
    const errorMessage = 'Error updating project: ' + (error instanceof Error ? error.message : String(error));
    Logger.write(errorMessage, LogLevel.Error);
    throw new Error(errorMessage);
  }
}

// Delete a project
export async function deleteProject(projectId: number): Promise<void> {
  try {
    Logger.write('Deleting project with ID: ' + projectId, LogLevel.Info);
    
    await sp.web.lists.getByTitle('Projects').items.getById(projectId).delete();
    
    Logger.write('Successfully deleted project with ID: ' + projectId, LogLevel.Info);
    
  } catch (error) {
    const errorMessage = 'Error deleting project: ' + (error instanceof Error ? error.message : String(error));
    Logger.write(errorMessage, LogLevel.Error);
    throw new Error(errorMessage);
  }
}

// Get project summary statistics
export async function getProjectSummary(): Promise<ProjectSummary> {
  try {
    Logger.write('Fetching project summary', LogLevel.Info);
    
    const projects = await getProjects();
    
    var totalProjects = projects.length;
    var activeProjects = 0;
    var completedProjects = 0;
    var onHoldProjects = 0;
    var notStartedProjects = 0;
    var totalProgress = 0;
    var totalBudget = 0;
    
    for (var i = 0; i < projects.length; i++) {
      var project = projects[i];
      
      switch (project.Status) {
        case Status.InProgress:
          activeProjects++;
          break;
        case Status.Completed:
          completedProjects++;
          break;
        case Status.OnHold:
          onHoldProjects++;
          break;
        case Status.NotStarted:
          notStartedProjects++;
          break;
      }
      
      totalProgress += project.Progress || 0;
      totalBudget += project.Budget || 0;
    }

    const summary: ProjectSummary = {
      totalProjects: totalProjects,
      activeProjects: activeProjects,
      completedProjects: completedProjects,
      onHoldProjects: onHoldProjects,
      notStartedProjects: notStartedProjects,
      averageProgress: totalProjects > 0 ? totalProgress / totalProjects : 0,
      totalBudget: totalBudget,
    };

    Logger.write('Successfully calculated project summary', LogLevel.Info);
    return summary;
    
  } catch (error) {
    const errorMessage = 'Error fetching project summary: ' + (error instanceof Error ? error.message : String(error));
    Logger.write(errorMessage, LogLevel.Error);
    throw new Error(errorMessage);
  }
}

// Bulk update projects
export async function bulkUpdateProjects(projectIds: number[], updates: Partial<Project>): Promise<Project[]> {
  try {
    Logger.write('Bulk updating ' + projectIds.length + ' projects', LogLevel.Info);
    
    const updatePromises = projectIds.map(function(projectId) {
      return getProject(projectId).then(function(project) {
        const updatedProject = Object.assign({}, project, updates, {
          ModifiedDate: new Date(),
        });
        return updateProject(updatedProject);
      });
    });

    const updatedProjects = await Promise.all(updatePromises);
    
    Logger.write('Successfully bulk updated ' + updatedProjects.length + ' projects', LogLevel.Info);
    return updatedProjects;
    
  } catch (error) {
    const errorMessage = 'Error bulk updating projects: ' + (error instanceof Error ? error.message : String(error));
    Logger.write(errorMessage, LogLevel.Error);
    throw new Error(errorMessage);
  }
}

// Utility function to ensure SharePoint list exists
export async function ensureProjectsList(): Promise<void> {
  try {
    Logger.write('Checking if Projects list exists', LogLevel.Info);
    
    // Try to get the list
    try {
      await sp.web.lists.getByTitle('Projects').get();
      Logger.write('Projects list already exists', LogLevel.Info);
      return;
    } catch (listError) {
      Logger.write('Projects list does not exist, creating...', LogLevel.Info);
    }

    // Create the list if it doesn't exist
    const listCreationInfo = {
      Title: 'Projects',
      Description: 'Project management list for SPFx template',
      TemplateFeatureId: '00bfea71-de22-43b2-a848-c05709900100', // Custom List template
      QuickLaunchOption: 1, // Show on Quick Launch
    };

    const list = await sp.web.lists.add(listCreationInfo.Title, listCreationInfo.Description, 100, false);

    // Add custom fields
    const fields = [
      { name: 'Status', type: 'Choice', choices: ['Not Started', 'In Progress', 'Completed', 'On Hold'] },
      { name: 'Priority', type: 'Number' },
      { name: 'StartDate', type: 'DateTime' },
      { name: 'EndDate', type: 'DateTime' },
      { name: 'Progress', type: 'Number' },
      { name: 'Budget', type: 'Currency' },
      { name: 'Owner', type: 'Text' },
      { name: 'Tags', type: 'Note' },
    ];

    for (var i = 0; i < fields.length; i++) {
      var field = fields[i];
      try {
        if (field.type === 'Choice') {
          await list.fields.addChoice(field.name, field.choices);
        } else if (field.type === 'Number') {
          await list.fields.addNumber(field.name);
        } else if (field.type === 'DateTime') {
          await list.fields.addDateTime(field.name);
        } else if (field.type === 'Currency') {
          await list.fields.addCurrency(field.name);
        } else if (field.type === 'Text') {
          await list.fields.addText(field.name);
        } else if (field.type === 'Note') {
          await list.fields.addMultilineText(field.name);
        }
      } catch (fieldError) {
        Logger.write('Error adding field ' + field.name + ': ' + (fieldError instanceof Error ? fieldError.message : String(fieldError)), LogLevel.Warning);
      }
    }

    Logger.write('Successfully created Projects list with custom fields', LogLevel.Info);
    
  } catch (error) {
    const errorMessage = 'Error ensuring Projects list: ' + (error instanceof Error ? error.message : String(error));
    Logger.write(errorMessage, LogLevel.Error);
    throw new Error(errorMessage);
  }
}
