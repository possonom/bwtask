import { useEffect, useState, useCallback } from 'react';

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { Logger, LogLevel } from "@pnp/logging";

import { getErrorMessage } from '../helper/errorHelper';
import { Project } from '../interfaces/Project';

const getData = async (): Promise<Project[]> => {
  Logger.write("getData - Start getting data...", LogLevel.Info);

  try {
    // In a real implementation, you would use getById for better performance
    const allItems: any[] = await sp.web.lists.getByTitle("Projects").items
      .select("Id", "Title", "Status")
      .top(100) // Limit results for performance
      .getAll();

    Logger.write(`getData - End getting data. Found ${allItems.length} items`, LogLevel.Info);

    const projects: Project[] = allItems.map(item => ({
      Id: item.Id,
      Name: item.Title,
      Status: item.Status || 0, // Default to 0 if Status is null/undefined
    }));
    
    return projects;
  } catch (error) {
    Logger.write(`getData - Error: ${getErrorMessage(error)}`, LogLevel.Error);
    throw error;
  }
};

// Mock data for development/testing when SharePoint is not available
const getMockData = (): Project[] => {
  return [
    { Id: 1, Name: "Website Redesign", Status: 1 },
    { Id: 2, Name: "Mobile App Development", Status: 2 },
    { Id: 3, Name: "Database Migration", Status: 0 },
    { Id: 4, Name: "Security Audit", Status: 3 },
    { Id: 5, Name: "Performance Optimization", Status: 1 },
    { Id: 6, Name: "User Training Program", Status: 2 },
    { Id: 7, Name: "API Integration", Status: 1 },
    { Id: 8, Name: "Documentation Update", Status: 0 },
  ];
};

export const useQueryProjects = () => {
  const [data, setData] = useState<Project[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isError, setIsError] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  const fetchData = useCallback(async () => {
    try {
      setIsLoading(true);
      setIsError(false);
      setError(null);
      
      let projects: Project[];
      
      try {
        // Try to get real data from SharePoint
        projects = await getData();
      } catch (spError) {
        // If SharePoint is not available, use mock data
        Logger.write("SharePoint not available, using mock data", LogLevel.Warning);
        projects = getMockData();
      }

      setData(projects);
    } catch (error: unknown) {
      const errMessage = getErrorMessage(error);
      Logger.write(`useQueryProjects - Error: ${errMessage}`, LogLevel.Error);
      
      setIsError(true);
      setError(errMessage);
      
      // Fallback to mock data even on error
      setData(getMockData());
    } finally {
      setIsLoading(false);
    } 
  }, []);

  useEffect(() => {
    fetchData();   
  }, [fetchData]);  

  return { 
    data, 
    isLoading, 
    isError, 
    error,
    refetch: fetchData
  };    
};
