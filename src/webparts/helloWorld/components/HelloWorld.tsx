import { useEffect } from 'react';
import { IHelloWorldProps } from './IHelloWorldProps';
import { useProjectsQuery } from '../../../hooks/useProjectData';
import { useProjectStore } from '../../../stores/projectStore';
import { Logger, LogLevel } from '@pnp/logging';

const HelloWorld = ({ description, context }: IHelloWorldProps) => {
  // Use React 17 with modern hooks for SharePoint 2019 compatibility
  const projectsQuery = useProjectsQuery();
  const { projects, isLoading, error } = useProjectStore();

  useEffect(() => {
    Logger.write('HelloWorld component mounted with React 17 for SharePoint 2019', LogLevel.Info);
    
    // Log React version for verification
    if (typeof window !== 'undefined' && (window as any).React) {
      Logger.write(`React version detected: ${(window as any).React.version}`, LogLevel.Info);
    }
  }, []);

  if (isLoading) {
    return (
      <div className="loading" style={{ padding: '20px', textAlign: 'center' }}>
        <div>Loading projects...</div>
        <div style={{ fontSize: '12px', color: '#666', marginTop: '10px' }}>
          React 17 + SPFx 1.4.1 + SharePoint 2019
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="error" style={{ padding: '20px', color: '#d13438' }}>
        <strong>Error:</strong> {error}
        <div style={{ fontSize: '12px', color: '#666', marginTop: '10px' }}>
          React 17 error handling in SharePoint 2019
        </div>
      </div>
    );
  }

  return (
    <div className="helloWorld" style={{ padding: '20px', fontFamily: 'Segoe UI, sans-serif' }}>
      <div style={{ 
        background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
        color: 'white',
        padding: '20px',
        borderRadius: '8px',
        marginBottom: '20px'
      }}>
        <h1 style={{ margin: '0 0 10px 0', fontSize: '24px' }}>
          🚀 React 17 + SharePoint 2019!
        </h1>
        <p style={{ margin: '0', opacity: 0.9 }}>{description}</p>
        <div style={{ 
          fontSize: '12px', 
          opacity: 0.8, 
          marginTop: '10px',
          padding: '8px 12px',
          background: 'rgba(255,255,255,0.1)',
          borderRadius: '4px',
          display: 'inline-block'
        }}>
          ✅ React 17.0.2 • SPFx 1.4.1 • SharePoint 2019 On-Premise
        </div>
      </div>

      <div style={{ 
        background: '#f8f9fa',
        border: '1px solid #e9ecef',
        borderRadius: '8px',
        padding: '20px',
        marginBottom: '20px'
      }}>
        <h2 style={{ margin: '0 0 15px 0', color: '#495057' }}>
          📊 Projects ({projects.length})
        </h2>
        
        {projects.length === 0 ? (
          <div style={{ 
            textAlign: 'center', 
            padding: '40px 20px',
            color: '#6c757d'
          }}>
            <div style={{ fontSize: '48px', marginBottom: '10px' }}>📝</div>
            <p>No projects found. Create your first project!</p>
            <div style={{ fontSize: '12px', marginTop: '10px' }}>
              React 17 state management with Zustand
            </div>
          </div>
        ) : (
          <div style={{ display: 'grid', gap: '10px' }}>
            {projects.map((project) => (
              <div 
                key={project.Id}
                style={{
                  background: 'white',
                  border: '1px solid #dee2e6',
                  borderRadius: '6px',
                  padding: '15px',
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center'
                }}
              >
                <div>
                  <strong style={{ color: '#212529' }}>{project.Name}</strong>
                  <div style={{ fontSize: '12px', color: '#6c757d', marginTop: '4px' }}>
                    Status: {project.Status} • Progress: {project.Progress || 0}%
                  </div>
                </div>
                <div style={{
                  background: project.Status === 'Completed' ? '#d4edda' : 
                           project.Status === 'In Progress' ? '#d1ecf1' : '#f8d7da',
                  color: project.Status === 'Completed' ? '#155724' : 
                         project.Status === 'In Progress' ? '#0c5460' : '#721c24',
                  padding: '4px 8px',
                  borderRadius: '4px',
                  fontSize: '12px',
                  fontWeight: 'bold'
                }}>
                  {project.Status}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      <div style={{ 
        background: '#e9ecef',
        borderRadius: '8px',
        padding: '20px'
      }}>
        <h3 style={{ margin: '0 0 15px 0', color: '#495057' }}>
          🔧 SharePoint Context
        </h3>
        <div style={{ display: 'grid', gap: '8px', fontSize: '14px' }}>
          <div>
            <strong>Site:</strong> {context.pageContext.web.title}
          </div>
          <div>
            <strong>User:</strong> {context.pageContext.user.displayName}
          </div>
          <div>
            <strong>Environment:</strong> SharePoint 2019 On-Premise
          </div>
          <div>
            <strong>SPFx Version:</strong> 1.4.1
          </div>
          <div style={{ 
            marginTop: '10px',
            padding: '10px',
            background: '#d4edda',
            borderRadius: '4px',
            color: '#155724'
          }}>
            <strong>🎉 React 17 Successfully Forced!</strong>
            <div style={{ fontSize: '12px', marginTop: '4px' }}>
              Modern React patterns running in legacy SharePoint environment
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default HelloWorld;
