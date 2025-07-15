import * as React from 'react';
import { useState, useCallback, useMemo } from 'react';
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
  Text,
  Card,
  ICardTokens,
  mergeStyles,
  FontWeights,
  SearchBox,
  ISearchBoxStyles,
  Separator,
  Icon,
  TooltipHost,
  DirectionalHint,
} from '@fluentui/react';

import { Project } from '../../../interfaces/Project';
import { useQueryProjects } from '../../../services/useQueryProjects';
import { Status } from '../../../interfaces/Status';
import { fluentTheme, themeTokens } from '../../../theme/fluentTheme';

// Styled components using Fluent UI 8 styling system
const containerClass = mergeStyles({
  padding: themeTokens.spacing.l,
  backgroundColor: fluentTheme.semanticColors.bodyBackground,
  minHeight: '100vh',
});

const headerClass = mergeStyles({
  marginBottom: themeTokens.spacing.l,
  padding: themeTokens.spacing.m,
  backgroundColor: fluentTheme.semanticColors.primaryButtonBackground,
  color: fluentTheme.semanticColors.primaryButtonText,
  borderRadius: themeTokens.borderRadius.medium,
  boxShadow: themeTokens.shadows.depth4,
});

const statsCardClass = mergeStyles({
  padding: themeTokens.spacing.m,
  backgroundColor: fluentTheme.semanticColors.bodyBackground,
  border: `1px solid ${fluentTheme.semanticColors.inputBorder}`,
  borderRadius: themeTokens.borderRadius.medium,
  boxShadow: themeTokens.shadows.depth4,
  marginBottom: themeTokens.spacing.m,
});

const SpDataAccess: React.FC<ISpDataAccessProps> = (props) => {
  const { description } = props;
  
  const stackTokens: IStackTokens = { childrenGap: 20 };
  const cardTokens: ICardTokens = { childrenMargin: 12 };
  
  const { data, isLoading, error, refetch } = useQueryProjects();
  
  const [selectedProject, setSelectedProject] = useState<Project | null>(null);
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [filterText, setFilterText] = useState<string>('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [showMessage, setShowMessage] = useState<boolean>(false);
  const [messageText, setMessageText] = useState<string>('');
  const [messageType, setMessageType] = useState<MessageBarType>(MessageBarType.info);

  // Filter projects based on search text and status
  const filteredProjects = useMemo(() => {
    if (!data) return [];
    
    return data.filter(project => {
      const matchesText = !filterText || 
        project.Name.toLowerCase().includes(filterText.toLowerCase());
      const matchesStatus = statusFilter === 'all' || 
        project.Status.toString() === statusFilter;
      
      return matchesText && matchesStatus;
    });
  }, [data, filterText, statusFilter]);

  // Project statistics
  const projectStats = useMemo(() => {
    if (!data) return { total: 0, completed: 0, inProgress: 0, notStarted: 0, onHold: 0 };
    
    return {
      total: data.length,
      completed: data.filter(p => p.Status === 2).length,
      inProgress: data.filter(p => p.Status === 1).length,
      notStarted: data.filter(p => p.Status === 0).length,
      onHold: data.filter(p => p.Status === 3).length,
    };
  }, [data]);

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

  const onSearchChange = useCallback((event?: React.ChangeEvent<HTMLInputElement>, newValue?: string) => {
    setFilterText(newValue || '');
  }, []);

  const statusOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Statuses', data: { icon: 'Filter' } },
    { key: '0', text: 'Not Started', data: { icon: 'CircleRing' } },
    { key: '1', text: 'In Progress', data: { icon: 'ProgressRingDots' } },
    { key: '2', text: 'Completed', data: { icon: 'CompletedSolid' } },
    { key: '3', text: 'On Hold', data: { icon: 'Pause' } },
  ];

  const onStatusFilterChange = useCallback((event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      setStatusFilter(option.key as string);
    }
  }, []);

  // Command bar items with modern Fluent UI 8 styling
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
    {
      key: 'new',
      text: 'New Project',
      iconProps: { iconName: 'Add' },
      onClick: () => showNotification('New project functionality would be implemented here'),
      buttonStyles: {
        root: {
          backgroundColor: fluentTheme.semanticColors.primaryButtonBackground,
          color: fluentTheme.semanticColors.primaryButtonText,
        },
      },
    },
  ];

  // Enhanced DetailsList columns with modern styling
  const columns: IColumn[] = [
    {
      key: 'id',
      name: 'ID',
      fieldName: 'Id',
      minWidth: 60,
      maxWidth: 80,
      isResizable: true,
      onRender: (item: Project) => (
        <Text variant="small" styles={{ root: { fontWeight: FontWeights.semibold } }}>
          #{item.Id}
        </Text>
      ),
    },
    {
      key: 'name',
      name: 'Project Name',
      fieldName: 'Name',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onRender: (item: Project) => (
        <TooltipHost content={`Click to view details for ${item.Name}`} directionalHint={DirectionalHint.topCenter}>
          <PrimaryButton
            text={item.Name}
            onClick={() => onProjectSelect(item)}
            styles={{
              root: {
                background: 'transparent',
                border: 'none',
                padding: 0,
                minWidth: 'auto',
                height: 'auto',
              },
              label: {
                fontWeight: FontWeights.regular,
                textDecoration: 'underline',
              },
            }}
          />
        </TooltipHost>
      ),
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 140,
      maxWidth: 180,
      isResizable: true,
      onRender: (item: Project) => {
        const statusText = Status[item.Status];
        const statusConfig = {
          0: { color: fluentTheme.semanticColors.bodySubtext, icon: 'CircleRing' }, // Not Started
          1: { color: fluentTheme.semanticColors.actionLink, icon: 'ProgressRingDots' }, // In Progress
          2: { color: fluentTheme.semanticColors.successText, icon: 'CompletedSolid' }, // Completed
          3: { color: fluentTheme.semanticColors.errorText, icon: 'Pause' }, // On Hold
        };
        
        const config = statusConfig[item.Status] || statusConfig[0];
        
        return (
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Icon iconName={config.icon} styles={{ root: { color: config.color } }} />
            <Text styles={{ root: { color: config.color, fontWeight: FontWeights.semibold } }}>
              {statusText}
            </Text>
          </Stack>
        );
      },
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 100,
      maxWidth: 120,
      onRender: (item: Project) => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <TooltipHost content="Edit project">
            <DefaultButton
              iconProps={{ iconName: 'Edit' }}
              onClick={() => showNotification(`Edit ${item.Name}`)}
              styles={{ root: { minWidth: 32, padding: '4px 8px' } }}
            />
          </TooltipHost>
          <TooltipHost content="More options">
            <DefaultButton
              iconProps={{ iconName: 'More' }}
              onClick={() => showNotification(`More options for ${item.Name}`)}
              styles={{ root: { minWidth: 32, padding: '4px 8px' } }}
            />
          </TooltipHost>
        </Stack>
      ),
    },
  ];

  const searchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 300,
    },
  };

  if (isLoading && !data) {
    return (
      <ThemeProvider theme={fluentTheme}>
        <div className={containerClass}>
          <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { height: '50vh' } }}>
            <Spinner size={SpinnerSize.large} label="Loading projects..." />
            <Text variant="medium" styles={{ root: { marginTop: themeTokens.spacing.m } }}>
              Fetching data from SharePoint...
            </Text>
          </Stack>
        </div>
      </ThemeProvider>
    );
  }

  if (error) {
    return (
      <ThemeProvider theme={fluentTheme}>
        <div className={containerClass}>
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            <Text variant="medium">
              <strong>Error loading data:</strong> {error}
            </Text>
          </MessageBar>
          <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: themeTokens.spacing.l } }}>
            <PrimaryButton text="Retry" iconProps={{ iconName: 'Refresh' }} onClick={onRefresh} />
          </Stack>
        </div>
      </ThemeProvider>
    );
  }

  return (
    <ThemeProvider theme={fluentTheme}>
      <div className={containerClass}>
        {/* Header Section */}
        <div className={headerClass}>
          <Text variant="xxLarge" styles={{ root: { fontWeight: FontWeights.bold, marginBottom: themeTokens.spacing.s } }}>
            📊 SharePoint Data Access
          </Text>
          <Text variant="medium" styles={{ root: { opacity: 0.9 } }}>
            {description}
          </Text>
          <Text variant="small" styles={{ root: { opacity: 0.8, marginTop: themeTokens.spacing.s } }}>
            Powered by PnP V2 • Fluent UI 8 • React 17
          </Text>
        </div>

        {/* Message Bar */}
        {showMessage && (
          <MessageBar
            messageBarType={messageType}
            isMultiline={false}
            onDismiss={() => setShowMessage(false)}
            dismissButtonAriaLabel="Close"
            styles={{ root: { marginBottom: themeTokens.spacing.m } }}
          >
            {messageText}
          </MessageBar>
        )}

        <Stack tokens={stackTokens}>
          {/* Statistics Cards */}
          <div className={statsCardClass}>
            <Text variant="large" styles={{ root: { fontWeight: FontWeights.semibold, marginBottom: themeTokens.spacing.m } }}>
              📈 Project Overview
            </Text>
            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
              <Stack.Item>
                <Stack horizontalAlign="center" tokens={{ childrenGap: 4 }}>
                  <Text variant="xxLarge" styles={{ root: { fontWeight: FontWeights.bold, color: fluentTheme.semanticColors.actionLink } }}>
                    {projectStats.total}
                  </Text>
                  <Text variant="small">Total Projects</Text>
                </Stack>
              </Stack.Item>
              <Stack.Item>
                <Stack horizontalAlign="center" tokens={{ childrenGap: 4 }}>
                  <Text variant="xxLarge" styles={{ root: { fontWeight: FontWeights.bold, color: fluentTheme.semanticColors.successText } }}>
                    {projectStats.completed}
                  </Text>
                  <Text variant="small">Completed</Text>
                </Stack>
              </Stack.Item>
              <Stack.Item>
                <Stack horizontalAlign="center" tokens={{ childrenGap: 4 }}>
                  <Text variant="xxLarge" styles={{ root: { fontWeight: FontWeights.bold, color: fluentTheme.semanticColors.actionLink } }}>
                    {projectStats.inProgress}
                  </Text>
                  <Text variant="small">In Progress</Text>
                </Stack>
              </Stack.Item>
              <Stack.Item>
                <Stack horizontalAlign="center" tokens={{ childrenGap: 4 }}>
                  <Text variant="xxLarge" styles={{ root: { fontWeight: FontWeights.bold, color: fluentTheme.semanticColors.bodySubtext } }}>
                    {projectStats.notStarted}
                  </Text>
                  <Text variant="small">Not Started</Text>
                </Stack>
              </Stack.Item>
              <Stack.Item>
                <Stack horizontalAlign="center" tokens={{ childrenGap: 4 }}>
                  <Text variant="xxLarge" styles={{ root: { fontWeight: FontWeights.bold, color: fluentTheme.semanticColors.errorText } }}>
                    {projectStats.onHold}
                  </Text>
                  <Text variant="small">On Hold</Text>
                </Stack>
              </Stack.Item>
            </Stack>
          </div>

          {/* Command Bar */}
          <CommandBar
            items={commandBarItems}
            ariaLabel="Project management commands"
            styles={{
              root: {
                backgroundColor: fluentTheme.semanticColors.bodyBackground,
                borderRadius: themeTokens.borderRadius.medium,
                boxShadow: themeTokens.shadows.depth4,
              },
            }}
          />

          {/* Filters Section */}
          <Card tokens={cardTokens}>
            <Stack horizontal wrap tokens={{ childrenGap: 20 }} verticalAlign="end">
              <SearchBox
                placeholder="Search projects..."
                value={filterText}
                onChange={onSearchChange}
                styles={searchBoxStyles}
                iconProps={{ iconName: 'Search' }}
              />
              <Dropdown
                label="Filter by Status"
                options={statusOptions}
                selectedKey={statusFilter}
                onChange={onStatusFilterChange}
                styles={{ root: { width: 180 } }}
                onRenderOption={(option) => (
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    {option?.data?.icon && <Icon iconName={option.data.icon} />}
                    <span>{option?.text}</span>
                  </Stack>
                )}
              />
              {isLoading && <Spinner size={SpinnerSize.medium} />}
            </Stack>
          </Card>

          {/* Projects List */}
          <Card tokens={cardTokens}>
            <Stack tokens={{ childrenGap: 16 }}>
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Text variant="large" styles={{ root: { fontWeight: FontWeights.semibold } }}>
                  🗂️ Projects ({filteredProjects.length})
                </Text>
                {filteredProjects.length !== data?.length && (
                  <Text variant="small" styles={{ root: { color: fluentTheme.semanticColors.bodySubtext } }}>
                    Showing {filteredProjects.length} of {data?.length} projects
                  </Text>
                )}
              </Stack>
              
              <Separator />
              
              {filteredProjects.length === 0 ? (
                <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }} styles={{ root: { padding: '40px 20px' } }}>
                  <Icon iconName="SearchIssue" styles={{ root: { fontSize: '48px', color: fluentTheme.semanticColors.bodySubtext } }} />
                  <Text variant="large" styles={{ root: { color: fluentTheme.semanticColors.bodySubtext } }}>
                    {data?.length === 0 ? 'No projects found in SharePoint list.' : 'No projects match the current filters.'}
                  </Text>
                  {data?.length === 0 && (
                    <PrimaryButton 
                      text="Create First Project" 
                      iconProps={{ iconName: 'Add' }}
                      onClick={() => showNotification('Create project functionality would be implemented here')}
                    />
                  )}
                </Stack>
              ) : (
                <DetailsList
                  items={filteredProjects}
                  columns={columns}
                  layoutMode={DetailsListLayoutMode.justified}
                  selectionMode={SelectionMode.none}
                  isHeaderVisible={true}
                  styles={{
                    root: {
                      borderRadius: themeTokens.borderRadius.medium,
                      overflow: 'hidden',
                    },
                  }}
                />
              )}
            </Stack>
          </Card>
        </Stack>

        {/* Project Details Panel */}
        <Panel
          headerText={selectedProject ? `Project: ${selectedProject.Name}` : 'Project Details'}
          isOpen={isPanelOpen}
          onDismiss={() => setIsPanelOpen(false)}
          type={PanelType.medium}
          closeButtonAriaLabel="Close"
          styles={{
            main: {
              backgroundColor: fluentTheme.semanticColors.bodyBackground,
            },
          }}
        >
          {selectedProject && (
            <Stack tokens={stackTokens}>
              <Card tokens={cardTokens}>
                <Text variant="large" styles={{ root: { fontWeight: FontWeights.semibold, marginBottom: themeTokens.spacing.m } }}>
                  📋 Project Information
                </Text>
                <Stack tokens={{ childrenGap: 12 }}>
                  <Stack horizontal horizontalAlign="space-between">
                    <Text variant="medium"><strong>ID:</strong></Text>
                    <Text variant="medium">#{selectedProject.Id}</Text>
                  </Stack>
                  <Stack horizontal horizontalAlign="space-between">
                    <Text variant="medium"><strong>Name:</strong></Text>
                    <Text variant="medium">{selectedProject.Name}</Text>
                  </Stack>
                  <Stack horizontal horizontalAlign="space-between">
                    <Text variant="medium"><strong>Status:</strong></Text>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Icon 
                        iconName={selectedProject.Status === 2 ? 'CompletedSolid' : selectedProject.Status === 1 ? 'ProgressRingDots' : 'CircleRing'} 
                        styles={{ root: { color: selectedProject.Status === 2 ? fluentTheme.semanticColors.successText : fluentTheme.semanticColors.actionLink } }}
                      />
                      <Text variant="medium">{Status[selectedProject.Status]}</Text>
                    </Stack>
                  </Stack>
                </Stack>
              </Card>
              
              <Card tokens={cardTokens}>
                <Text variant="large" styles={{ root: { fontWeight: FontWeights.semibold, marginBottom: themeTokens.spacing.m } }}>
                  ⚡ Quick Actions
                </Text>
                <Stack horizontal tokens={{ childrenGap: 12 }}>
                  <PrimaryButton 
                    text="Edit Project" 
                    iconProps={{ iconName: 'Edit' }}
                    onClick={() => showNotification('Edit functionality would be implemented here')}
                  />
                  <DefaultButton 
                    text="View Timeline" 
                    iconProps={{ iconName: 'Timeline' }}
                    onClick={() => showNotification('Timeline view would be implemented here')}
                  />
                  <DefaultButton 
                    text="Export Data" 
                    iconProps={{ iconName: 'Download' }}
                    onClick={() => showNotification('Export functionality would be implemented here')}
                  />
                </Stack>
              </Card>
            </Stack>
          )}
        </Panel>
      </div>
    </ThemeProvider>
  );
};

export default SpDataAccess;
