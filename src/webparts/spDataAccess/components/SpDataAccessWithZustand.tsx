/**
 * Enhanced SpDataAccess component using Zustand and TanStack Query
 */

import * as React from 'react';
import { useEffect, useCallback } from 'react';
import { ISpDataAccessProps } from './ISpDataAccessProps';
import {
  ThemeProvider,
  Stack,
  IStackTokens,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  TextField,
  Dropdown,
  IDropdownOption,
  Panel,
  PanelType,
  CommandBar,
  ICommandBarItemProps,
  IconButton,
  TooltipHost,
  Separator,
  Text,
} from '@fluentui/react';

import { QueryProvider } from '../../../providers/QueryProvider';
import { useProjectSelectors, useProjectActions } from '../../../stores/projectStore';
import { useProjects, useCreateProject, useUpdateProject, useDeleteProject } from '../../../hooks/useProjectQueries';
import { Project, ProjectCreateRequest, ProjectUpdateRequest } from '../../../interfaces/Project';
import { Status, getStatusInfo, getStatusOptions } from '../../../interfaces/Status';

const stackTokens: IStackTokens = { childrenGap: 20 };

const SpDataAccessContent: React.FC<ISpDataAccessProps> = (props) => {
  const { theme, description } = props;
  
  // Zustand selectors
  const {
    projects,
    selectedProject,
    isLoading,
    error,
    summary,
    filters,
    searchText,
  } = useProjectSelectors();
  
  // Zustand actions
  const {
    setSearchText,
    setFilters,
    setSorting,
    openPanel,
    closePanel,
    setError,
  } = useProjectActions();
  
  // TanStack Query hooks
  const { refetch, isFetching } = useProjects(filters);
  const createProjectMutation = useCreateProject();
  const updateProjectMutation = useUpdateProject();
  const deleteProjectMutation = useDeleteProject();
  
  // Local state for form
  const [formData, setFormData] = React.useState<Partial<Project>>({});
  const [showMessage, setShowMessage] = React.useState<boolean>(false);
  const [messageText, setMessageText] = React.useState<string>('');
  const [messageType, setMessageType] = React.useState<MessageBarType>(MessageBarType.info);

  // Initialize data on mount
  useEffect(() => {
    refetch();
  }, [refetch]);

  const showNotification = useCallback((message: string, type: MessageBarType = MessageBarType.info) => {
    setMessageText(message);
    setMessageType(type);
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 4000);
  }, []);

  const onRefresh = useCallback(async () => {
    try {
      await refetch();
      showNotification('Data refreshed successfully!', MessageBarType.success);
    } catch (err) {
      showNotification('Failed to refresh data', MessageBarType.error);
    }
  }, [refetch, showNotification]);

  const onProjectSelect = useCallback((project: Project) => {
    setFormData(project);
    openPanel(project);
  }, [openPanel]);

  const onCreateNew = useCallback(() => {
    setFormData({
      Name: '',
      Status: Status.NotStarted,
      Description: '',
      Priority: 0,
      Progress: 0,
    });
    openPanel();
  }, [openPanel]);

  const onSaveProject = useCallback(async () => {
    if (!formData.Name) {
      showNotification('Project name is required', MessageBarType.error);
      return;
    }

    try {
      if (selectedProject) {
        // Update existing project
        const updateData: ProjectUpdateRequest = {
          Id: selectedProject.Id,
          ...formData,
        };
        await updateProjectMutation.mutateAsync(updateData);
        showNotification('Project updated successfully!', MessageBarType.success);
      } else {
        // Create new project
        const createData: ProjectCreateRequest = {
          Name: formData.Name!,
          Status: formData.Status || Status.NotStarted,
          Description: formData.Description,
          Priority: formData.Priority,
        };
        await createProjectMutation.mutateAsync(createData);
        showNotification('Project created successfully!', MessageBarType.success);
      }
      
      closePanel();
      setFormData({});
    } catch (error) {
      showNotification('Failed to save project', MessageBarType.error);
    }
  }, [formData, selectedProject, updateProjectMutation, createProjectMutation, closePanel, showNotification]);

  const onDeleteProject = useCallback(async (projectId: number) => {
    if (window.confirm('Are you sure you want to delete this project?')) {
      try {
        await deleteProjectMutation.mutateAsync(projectId);
        showNotification('Project deleted successfully!', MessageBarType.success);
        closePanel();
      } catch (error) {
        showNotification('Failed to delete project', MessageBarType.error);
      }
    }
  }, [deleteProjectMutation, closePanel, showNotification]);

  const onSearchTextChange = useCallback((event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    setSearchText(newValue || '');
  }, [setSearchText]);

  const statusOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Statuses' },
    ...getStatusOptions(),
  ];

  const onStatusFilterChange = useCallback((event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      const statusValue = option.key === 'all' ? undefined : parseInt(option.key as string);
      setFilters({ status: statusValue });
    }
  }, [setFilters]);

  const onColumnClick = useCallback((column: IColumn) => {
    const sortBy = column.key as any;
    setSorting(sortBy, 'asc'); // You could toggle direction here
  }, [setSorting]);

  // Command bar items
  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'new',
      text: 'New Project',
      iconProps: { iconName: 'Add' },
      onClick: onCreateNew,
    },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: onRefresh,
      disabled: isLoading || isFetching,
    },
    {
      key: 'export',
      text: 'Export',
      iconProps: { iconName: 'Download' },
      onClick: () => showNotification('Export functionality would be implemented here'),
    },
  ];

  // DetailsList columns
  const columns: IColumn[] = [
    {
      key: 'id',
      name: 'ID',
      fieldName: 'Id',
      minWidth: 50,
      maxWidth: 80,
      isResizable: true,
      onColumnClick,
    },
    {
      key: 'name',
      name: 'Project Name',
      fieldName: 'Name',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onColumnClick,
      onRender: (item: Project) => (
        <button
          style={{
            background: 'none',
            border: 'none',
            color: '#0078d4',
            cursor: 'pointer',
            textDecoration: 'underline',
            padding: 0,
            textAlign: 'left',
          }}
          onClick={() => onProjectSelect(item)}
        >
          {item.Name}
        </button>
      ),
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onColumnClick,
      onRender: (item: Project) => {
        const statusInfo = getStatusInfo(item.Status);
        return (
          <span style={{ 
            color: statusInfo.color, 
            fontWeight: 'semibold',
            padding: '4px 8px',
            borderRadius: '4px',
            backgroundColor: statusInfo.backgroundColor,
          }}>
            {statusInfo.label}
          </span>
        );
      },
    },
    {
      key: 'progress',
      name: 'Progress',
      fieldName: 'Progress',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onColumnClick,
      onRender: (item: Project) => (
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <div style={{
            width: '60px',
            height: '8px',
            backgroundColor: '#f3f2f1',
            borderRadius: '4px',
            overflow: 'hidden',
          }}>
            <div style={{
              width: `${item.Progress || 0}%`,
              height: '100%',
              backgroundColor: '#0078d4',
              transition: 'width 0.3s ease',
            }} />
          </div>
          <Text variant="small">{item.Progress || 0}%</Text>
        </div>
      ),
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 80,
      maxWidth: 100,
      onRender: (item: Project) => (
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          <TooltipHost content="Edit project">
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              onClick={() => onProjectSelect(item)}
            />
          </TooltipHost>
          <TooltipHost content="Delete project">
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              onClick={() => onDeleteProject(item.Id)}
            />
          </TooltipHost>
        </Stack>
      ),
    },
  ];

  if (isLoading && projects.length === 0) {
    return (
      <ThemeProvider theme={theme}>
        <div style={{ padding: '20px', textAlign: 'center' }}>
          <Spinner size={SpinnerSize.large} label="Loading projects..." />
        </div>
      </ThemeProvider>
    );
  }

  return (
    <ThemeProvider theme={theme}>
      <div style={{ padding: '20px' }}>
        <h1>SharePoint Data Access with Zustand & TanStack Query</h1>
        <p>{description}</p>
        <p>This component demonstrates modern state management with Zustand and data fetching with TanStack Query.</p>

        {/* Message Bar */}
        {showMessage && (
          <MessageBar
            messageBarType={messageType}
            isMultiline={false}
            onDismiss={() => setShowMessage(false)}
            dismissButtonAriaLabel="Close"
          >
            {messageText}
          </MessageBar>
        )}

        {/* Error Display */}
        {error && (
          <MessageBar messageBarType={MessageBarType.error}>
            <strong>Error:</strong> {error}
          </MessageBar>
        )}

        <Stack tokens={stackTokens}>
          {/* Command Bar */}
          <CommandBar
            items={commandBarItems}
            ariaLabel="Project management commands"
          />

          {/* Filters */}
          <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="end">
            <TextField
              label="Search Projects"
              placeholder="Type to filter projects..."
              value={searchText}
              onChange={onSearchTextChange}
              styles={{ root: { width: 250 } }}
            />
            <Dropdown
              label="Filter by Status"
              options={statusOptions}
              selectedKey={filters.status?.toString() || 'all'}
              onChange={onStatusFilterChange}
              styles={{ root: { width: 150 } }}
            />
            {(isLoading || isFetching) && <Spinner size={SpinnerSize.medium} />}
          </Stack>

          {/* Statistics */}
          {summary && (
            <div style={{ 
              padding: '15px', 
              backgroundColor: '#f3f2f1', 
              borderRadius: '4px',
              display: 'flex',
              gap: '30px',
              flexWrap: 'wrap',
            }}>
              <div><strong>Total:</strong> {summary.totalProjects}</div>
              <div><strong>Active:</strong> {summary.activeProjects}</div>
              <div><strong>Completed:</strong> {summary.completedProjects}</div>
              <div><strong>On Hold:</strong> {summary.onHoldProjects}</div>
              <div><strong>Avg Progress:</strong> {Math.round(summary.averageProgress)}%</div>
              {summary.totalBudget > 0 && (
                <div><strong>Total Budget:</strong> ${summary.totalBudget.toLocaleString()}</div>
              )}
            </div>
          )}

          {/* Projects List */}
          <div>
            <h3>Projects ({projects.length})</h3>
            {projects.length === 0 ? (
              <div style={{ textAlign: 'center', padding: '40px', color: '#605e5c' }}>
                No projects found. Click "New Project" to create one.
              </div>
            ) : (
              <DetailsList
                items={projects}
                columns={columns}
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                isHeaderVisible={true}
              />
            )}
          </div>
        </Stack>

        {/* Project Form Panel */}
        <Panel
          headerText={selectedProject ? `Edit: ${selectedProject.Name}` : 'Create New Project'}
          isOpen={useProjectSelectors().projects.length >= 0 && React.useMemo(() => {
            return useProjectStore.getState().isPanelOpen;
          }, [useProjectStore.getState().isPanelOpen])}
          onDismiss={closePanel}
          type={PanelType.medium}
          closeButtonAriaLabel="Close"
        >
          <Stack tokens={stackTokens}>
            <TextField
              label="Project Name *"
              value={formData.Name || ''}
              onChange={(_, value) => setFormData(prev => ({ ...prev, Name: value }))}
              required
            />
            
            <TextField
              label="Description"
              multiline
              rows={3}
              value={formData.Description || ''}
              onChange={(_, value) => setFormData(prev => ({ ...prev, Description: value }))}
            />
            
            <Dropdown
              label="Status"
              options={getStatusOptions()}
              selectedKey={formData.Status?.toString()}
              onChange={(_, option) => setFormData(prev => ({ 
                ...prev, 
                Status: option ? parseInt(option.key as string) : Status.NotStarted 
              }))}
            />
            
            <TextField
              label="Progress (%)"
              type="number"
              min={0}
              max={100}
              value={formData.Progress?.toString() || '0'}
              onChange={(_, value) => setFormData(prev => ({ 
                ...prev, 
                Progress: parseInt(value || '0') 
              }))}
            />

            <Separator />
            
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton 
                text={selectedProject ? 'Update' : 'Create'} 
                onClick={onSaveProject}
                disabled={createProjectMutation.isLoading || updateProjectMutation.isLoading}
              />
              <DefaultButton 
                text="Cancel" 
                onClick={closePanel} 
              />
              {selectedProject && (
                <DefaultButton 
                  text="Delete" 
                  onClick={() => onDeleteProject(selectedProject.Id)}
                  disabled={deleteProjectMutation.isLoading}
                  styles={{ root: { marginLeft: 'auto' } }}
                />
              )}
            </Stack>
          </Stack>
        </Panel>
      </div>
    </ThemeProvider>
  );
};

// Wrapper component with QueryProvider
const SpDataAccessWithZustand: React.FC<ISpDataAccessProps> = (props) => {
  return (
    <QueryProvider>
      <SpDataAccessContent {...props} />
    </QueryProvider>
  );
};

export default SpDataAccessWithZustand;
