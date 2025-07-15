import { useEffect } from 'react';
import { IHelloWorldProps } from './IHelloWorldProps';
import { useProjectsQuery } from '../../../hooks/useProjectData';
import { useProjectStore } from '../../../stores/projectStore';
import { Logger, LogLevel } from '@pnp/logging';
import {
  ThemeProvider,
  Stack,
  Text,
  Card,
  ICardTokens,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Icon,
  mergeStyles,
  FontWeights,
  Separator,
  ProgressIndicator,
  PersonaCoin,
  PersonaSize,
} from '@fluentui/react';
import { fluentTheme, themeTokens } from '../../../theme/fluentTheme';

// Styled components using Fluent UI 8 styling system
const containerClass = mergeStyles({
  padding: themeTokens.spacing.l,
  backgroundColor: fluentTheme.semanticColors.bodyBackground,
  minHeight: '100vh',
});

const heroClass = mergeStyles({
  background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
  color: fluentTheme.semanticColors.primaryButtonText,
  padding: themeTokens.spacing.xl,
  borderRadius: themeTokens.borderRadius.large,
  marginBottom: themeTokens.spacing.l,
  boxShadow: themeTokens.shadows.depth16,
  position: 'relative',
  overflow: 'hidden',
  '&::before': {
    content: '""',
    position: 'absolute',
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    background: 'rgba(255, 255, 255, 0.1)',
    backdropFilter: 'blur(10px)',
    zIndex: 0,
  },
});

const projectCardClass = mergeStyles({
  backgroundColor: fluentTheme.semanticColors.bodyBackground,
  border: `1px solid ${fluentTheme.semanticColors.inputBorder}`,
  borderRadius: themeTokens.borderRadius.medium,
  padding: themeTokens.spacing.m,
  marginBottom: themeTokens.spacing.s,
  boxShadow: themeTokens.shadows.depth4,
  transition: 'all 0.2s ease-in-out',
  '&:hover': {
    boxShadow: themeTokens.shadows.depth8,
    transform: 'translateY(-2px)',
  },
});

const HelloWorld = ({ description, context }: IHelloWorldProps) => {
  const projectsQuery = useProjectsQuery();
  const { projects, isLoading, error } = useProjectStore();
  const cardTokens: ICardTokens = { childrenMargin: 16 };

  useEffect(() => {
    Logger.write('HelloWorld component mounted with React 17 for SharePoint 2019', LogLevel.Info);
    
    if (typeof window !== 'undefined' && (window as any).React) {
      Logger.write(`React version detected: ${(window as any).React.version}`, LogLevel.Info);
    }
  }, []);

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'Completed':
        return fluentTheme.semanticColors.successText;
      case 'In Progress':
        return fluentTheme.semanticColors.actionLink;
      case 'On Hold':
        return fluentTheme.semanticColors.errorText;
      default:
        return fluentTheme.semanticColors.bodySubtext;
    }
  };

  const getStatusIcon = (status: string) => {
    switch (status) {
      case 'Completed':
        return 'CompletedSolid';
      case 'In Progress':
        return 'ProgressRingDots';
      case 'On Hold':
        return 'Pause';
      default:
        return 'CircleRing';
    }
  };

  const calculateProgress = (status: string) => {
    switch (status) {
      case 'Completed':
        return 1;
      case 'In Progress':
        return 0.6;
      case 'On Hold':
        return 0.3;
      default:
        return 0;
    }
  };

  if (isLoading) {
    return (
      <ThemeProvider theme={fluentTheme}>
        <div className={containerClass}>
          <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { height: '50vh' } }}>
            <Spinner size={SpinnerSize.large} label="Loading projects..." />
            <Text variant="medium" styles={{ root: { marginTop: themeTokens.spacing.m, color: fluentTheme.semanticColors.bodySubtext } }}>
              React 17 + SPFx 1.4.1 + SharePoint 2019
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
              <strong>Error:</strong> {error}
            </Text>
            <Text variant="small" styles={{ root: { marginTop: themeTokens.spacing.s } }}>
              React 17 error handling in SharePoint 2019
            </Text>
          </MessageBar>
        </div>
      </ThemeProvider>
    );
  }

  return (
    <ThemeProvider theme={fluentTheme}>
      <div className={containerClass}>
        {/* Hero Section */}
        <div className={heroClass} style={{ position: 'relative', zIndex: 1 }}>
          <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16 }}>
              <Icon iconName="Rocket" styles={{ root: { fontSize: '32px' } }} />
              <Text variant="xxLarge" styles={{ root: { fontWeight: FontWeights.bold } }}>
                React 17 + SharePoint 2019!
              </Text>
            </Stack>
            <Text variant="large" styles={{ root: { opacity: 0.9 } }}>
              {description}
            </Text>
            <div style={{ 
              backgroundColor: 'rgba(255,255,255,0.15)',
              borderRadius: themeTokens.borderRadius.medium,
              padding: themeTokens.spacing.m,
              display: 'inline-block',
              backdropFilter: 'blur(10px)',
            }}>
              <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                <Icon iconName="CheckMark" styles={{ root: { color: '#4CAF50' } }} />
                <Text variant="medium" styles={{ root: { fontWeight: FontWeights.semibold } }}>
                  React 17.0.2 • SPFx 1.4.1 • SharePoint 2019 On-Premise
                </Text>
              </Stack>
            </div>
          </Stack>
        </div>

        <Stack tokens={{ childrenGap: 24 }}>
          {/* Projects Overview Card */}
          <Card tokens={cardTokens}>
            <Stack tokens={{ childrenGap: 16 }}>
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                  <Icon iconName="ProjectCollection" styles={{ root: { fontSize: '24px', color: fluentTheme.semanticColors.actionLink } }} />
                  <Text variant="xLarge" styles={{ root: { fontWeight: FontWeights.semibold } }}>
                    Projects ({projects.length})
                  </Text>
                </Stack>
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  <DefaultButton 
                    text="Refresh" 
                    iconProps={{ iconName: 'Refresh' }}
                    onClick={() => window.location.reload()}
                  />
                  <PrimaryButton 
                    text="New Project" 
                    iconProps={{ iconName: 'Add' }}
                    onClick={() => alert('Create project functionality would be implemented here')}
                  />
                </Stack>
              </Stack>
              
              <Separator />
              
              {projects.length === 0 ? (
                <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }} styles={{ root: { padding: '40px 20px' } }}>
                  <Icon iconName="DocumentAdd" styles={{ root: { fontSize: '64px', color: fluentTheme.semanticColors.bodySubtext } }} />
                  <Text variant="large" styles={{ root: { color: fluentTheme.semanticColors.bodySubtext } }}>
                    No projects found. Create your first project!
                  </Text>
                  <Text variant="small" styles={{ root: { color: fluentTheme.semanticColors.bodySubtext } }}>
                    React 17 state management with Zustand
                  </Text>
                  <PrimaryButton 
                    text="Get Started" 
                    iconProps={{ iconName: 'Add' }}
                    onClick={() => alert('Create project functionality would be implemented here')}
                  />
                </Stack>
              ) : (
                <Stack tokens={{ childrenGap: 12 }}>
                  {projects.map((project) => (
                    <div key={project.Id} className={projectCardClass}>
                      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { flex: 1 } }}>
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                            <Text variant="mediumPlus" styles={{ root: { fontWeight: FontWeights.semibold } }}>
                              {project.Name}
                            </Text>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                              <Icon 
                                iconName={getStatusIcon(project.Status)} 
                                styles={{ root: { color: getStatusColor(project.Status), fontSize: '16px' } }}
                              />
                              <Text 
                                variant="small" 
                                styles={{ root: { 
                                  color: getStatusColor(project.Status), 
                                  fontWeight: FontWeights.semibold,
                                  textTransform: 'uppercase',
                                  letterSpacing: '0.5px'
                                } }}
                              >
                                {project.Status}
                              </Text>
                            </Stack>
                          </Stack>
                          
                          <ProgressIndicator 
                            percentComplete={calculateProgress(project.Status)}
                            styles={{
                              root: { width: '200px' },
                              progressBar: {
                                backgroundColor: getStatusColor(project.Status),
                              },
                            }}
                          />
                          
                          <Text variant="small" styles={{ root: { color: fluentTheme.semanticColors.bodySubtext } }}>
                            Progress: {Math.round(calculateProgress(project.Status) * 100)}% • ID: #{project.Id}
                          </Text>
                        </Stack>
                        
                        <Stack horizontal tokens={{ childrenGap: 8 }}>
                          <DefaultButton 
                            iconProps={{ iconName: 'View' }}
                            title="View Details"
                            onClick={() => alert(`View details for ${project.Name}`)}
                            styles={{ root: { minWidth: 32 } }}
                          />
                          <DefaultButton 
                            iconProps={{ iconName: 'Edit' }}
                            title="Edit Project"
                            onClick={() => alert(`Edit ${project.Name}`)}
                            styles={{ root: { minWidth: 32 } }}
                          />
                          <DefaultButton 
                            iconProps={{ iconName: 'More' }}
                            title="More Options"
                            onClick={() => alert(`More options for ${project.Name}`)}
                            styles={{ root: { minWidth: 32 } }}
                          />
                        </Stack>
                      </Stack>
                    </div>
                  ))}
                </Stack>
              )}
            </Stack>
          </Card>

          {/* SharePoint Context Card */}
          <Card tokens={cardTokens}>
            <Stack tokens={{ childrenGap: 16 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                <Icon iconName="SharePointLogo" styles={{ root: { fontSize: '24px', color: fluentTheme.semanticColors.actionLink } }} />
                <Text variant="xLarge" styles={{ root: { fontWeight: FontWeights.semibold } }}>
                  SharePoint Context
                </Text>
              </Stack>
              
              <Separator />
              
              <Stack tokens={{ childrenGap: 16 }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                    <Icon iconName="Globe" styles={{ root: { color: fluentTheme.semanticColors.actionLink } }} />
                    <Text variant="medium"><strong>Site:</strong></Text>
                  </Stack>
                  <Text variant="medium">{context.pageContext.web.title}</Text>
                </Stack>
                
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                    <PersonaCoin size={PersonaSize.size24} text={context.pageContext.user.displayName} />
                    <Text variant="medium"><strong>User:</strong></Text>
                  </Stack>
                  <Text variant="medium">{context.pageContext.user.displayName}</Text>
                </Stack>
                
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                    <Icon iconName="Server" styles={{ root: { color: fluentTheme.semanticColors.actionLink } }} />
                    <Text variant="medium"><strong>Environment:</strong></Text>
                  </Stack>
                  <Text variant="medium">SharePoint 2019 On-Premise</Text>
                </Stack>
                
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                    <Icon iconName="Code" styles={{ root: { color: fluentTheme.semanticColors.actionLink } }} />
                    <Text variant="medium"><strong>SPFx Version:</strong></Text>
                  </Stack>
                  <Text variant="medium">1.4.1</Text>
                </Stack>
              </Stack>
              
              <div style={{ 
                marginTop: themeTokens.spacing.m,
                padding: themeTokens.spacing.m,
                backgroundColor: '#d4edda',
                borderRadius: themeTokens.borderRadius.medium,
                border: '1px solid #c3e6cb'
              }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                  <Icon iconName="PartyLeader" styles={{ root: { color: '#155724', fontSize: '20px' } }} />
                  <Stack>
                    <Text variant="medium" styles={{ root: { color: '#155724', fontWeight: FontWeights.bold } }}>
                      🎉 React 17 Successfully Forced!
                    </Text>
                    <Text variant="small" styles={{ root: { color: '#155724' } }}>
                      Modern React patterns running in legacy SharePoint environment
                    </Text>
                  </Stack>
                </Stack>
              </div>
            </Stack>
          </Card>
        </Stack>
      </div>
    </ThemeProvider>
  );
};

export default HelloWorld;
