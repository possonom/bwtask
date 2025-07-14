import * as React from 'react';
import { useState, useCallback } from 'react';
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
} from '@fluentui/react';

import { Project } from '../../../interfaces/Project';
import { useQueryProjects } from '../../../services/useQueryProjects';
import { Status } from '../../../interfaces/Status';

const SpDataAccess: React.FC<ISpDataAccessProps> = (props) => {
  const { theme, description } = props;
  
  const stackTokens: IStackTokens = { childrenGap: 20 };
  const { data, isLoading, error, refetch } = useQueryProjects();
  
  const [selectedProject, setSelectedProject] = useState<Project | null>(null);
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [filterText, setFilterText] = useState<string>('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [showMessage, setShowMessage] = useState<boolean>(false);
  const [messageText, setMessageText] = useState<string>('');
  const [messageType, setMessageType] = useState<MessageBarType>(MessageBarType.info);

  // Filter projects based on search text and status
  const filteredProjects = React.useMemo(() => {
    if (!data) return [];
    
    return data.filter(project => {
      const matchesText = !filterText || 
        project.Name.toLowerCase().includes(filterText.toLowerCase());
      const matchesStatus = statusFilter === 'all' || 
        project.Status.toString() === statusFilter;
      
      return matchesText && matchesStatus;
    });
  }, [data, filterText, statusFilter]);

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
    setSelectedProject(project);
    setIsPanelOpen(true);
  }, []);

  const onFilterTextChange = useCallback((event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    setFilterText(newValue || '');
  }, []);

  const statusOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Statuses' },
    { key: '0', text: 'Not Started' },
    { key: '1', text: 'In Progress' },
    { key: '2', text: 'Completed' },
    { key: '3', text: 'On Hold' },
  ];

  const onStatusFilterChange = useCallback((event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      setStatusFilter(option.key as string);
    }
  }, []);

  // Command bar items
  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: onRefresh,
      disabled: isLoading,
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
    },
    {
      key: 'name',
      name: 'Project Name',
      fieldName: 'Name',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
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
      onRender: (item: Project) => {
        const statusText = Status[item.Status];
        const statusColor = 
          item.Status === 2 ? '#107c10' : // Completed - green
          item.Status === 1 ? '#0078d4' : // In Progress - blue
          item.Status === 3 ? '#d13438' : // On Hold - red
          '#605e5c'; // Not Started - gray
        
        return (
          <span style={{ color: statusColor, fontWeight: 'semibold' }}>
            {statusText}
          </span>
        );
      },
    },
  ];

  if (isLoading && !data) {
    return (
      <ThemeProvider theme={theme}>
        <div style={{ padding: '20px', textAlign: 'center' }}>
          <Spinner size={SpinnerSize.large} label="Loading projects..." />
        </div>
      </ThemeProvider>
    );
  }

  if (error) {
    return (
      <ThemeProvider theme={theme}>
        <div style={{ padding: '20px' }}>
          <MessageBar messageBarType={MessageBarType.error}>
            <strong>Error loading data:</strong> {error}
          </MessageBar>
          <br />
          <PrimaryButton text="Retry" onClick={onRefresh} />
        </div>
      </ThemeProvider>
    );
  }

  return (
    <ThemeProvider theme={theme}>
      <div style={{ padding: '20px' }}>
        <h1>SharePoint Data Access with PnP V2</h1>
        <p>{description}</p>
        <p>This component demonstrates how to fetch and display data from SharePoint using PnP V2 and custom hooks.</p>

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
              value={filterText}
              onChange={onFilterTextChange}
              styles={{ root: { width: 250 } }}
            />
            <Dropdown
              label="Filter by Status"
              options={statusOptions}
              selectedKey={statusFilter}
              onChange={onStatusFilterChange}
              styles={{ root: { width: 150 } }}
            />
            {isLoading && <Spinner size={SpinnerSize.medium} />}
          </Stack>

          {/* Statistics */}
          <div style={{ 
            padding: '15px', 
            backgroundColor: '#f3f2f1', 
            borderRadius: '4px',
            display: 'flex',
            gap: '30px'
          }}>
            <div><strong>Total Projects:</strong> {data?.length || 0}</div>
            <div><strong>Filtered Results:</strong> {filteredProjects.length}</div>
            <div><strong>Completed:</strong> {data?.filter(p => p.Status === 2).length || 0}</div>
            <div><strong>In Progress:</strong> {data?.filter(p => p.Status === 1).length || 0}</div>
          </div>

          {/* Projects List */}
          <div>
            <h3>Projects ({filteredProjects.length})</h3>
            {filteredProjects.length === 0 ? (
              <div style={{ textAlign: 'center', padding: '40px', color: '#605e5c' }}>
                {data?.length === 0 ? 'No projects found in SharePoint list.' : 'No projects match the current filters.'}
              </div>
            ) : (
              <DetailsList
                items={filteredProjects}
                columns={columns}
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                isHeaderVisible={true}
              />
            )}
          </div>
        </Stack>

        {/* Project Details Panel */}
        <Panel
          headerText={selectedProject ? `Project: ${selectedProject.Name}` : 'Project Details'}
          isOpen={isPanelOpen}
          onDismiss={() => setIsPanelOpen(false)}
          type={PanelType.medium}
          closeButtonAriaLabel="Close"
        >
          {selectedProject && (
            <Stack tokens={stackTokens}>
              <div>
                <h4>Project Information</h4>
                <div style={{ marginBottom: '10px' }}>
                  <strong>ID:</strong> {selectedProject.Id}
                </div>
                <div style={{ marginBottom: '10px' }}>
                  <strong>Name:</strong> {selectedProject.Name}
                </div>
                <div style={{ marginBottom: '10px' }}>
                  <strong>Status:</strong> {Status[selectedProject.Status]} ({selectedProject.Status})
                </div>
              </div>
              
              <div>
                <h4>Actions</h4>
                <Stack horizontal tokens={{ childrenGap: 10 }}>
                  <PrimaryButton 
                    text="Edit Project" 
                    onClick={() => showNotification('Edit functionality would be implemented here')}
                  />
                  <DefaultButton 
                    text="View Details" 
                    onClick={() => showNotification('View details functionality would be implemented here')}
                  />
                </Stack>
              </div>
            </Stack>
          )}
        </Panel>
      </div>
    </ThemeProvider>
  );
};

export default SpDataAccess;
