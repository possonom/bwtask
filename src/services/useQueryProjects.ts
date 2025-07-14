import { useEffect, useState } from 'react';

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { Logger, LogLevel } from "@pnp/logging";

import { getErrorMessage } from '../helper/errorHelper';

import { Project } from '../interfaces/Project';

const getData = async () => {

  Logger.write("getData - Start getting data...", LogLevel.Info);

  const allItems: any[] = await sp.web.lists.getByTitle("Projects").items.getAll(); // demo - use getById in real code!

  Logger.write(`getData - End getting data. Found ${allItems.length} items`, LogLevel.Info);

  const projects: Project[] = allItems.map(i => {
    return { 
        Id: i.Id,
        Name: i.Title,
        Status: i.Status,
    };
  });
  
  return projects;
};

export const useQueryProjects = () => {

  const [data, setData] = useState<Project[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isError, setIsError] = useState<boolean>(false);
  const [error, setError] = useState<string>(null);

  useEffect(() => {
      
    const execute = async () => {

      try {
        setIsLoading(true);
        
        const projects: Project[] = await getData();

        setData(projects);
        setIsError(false);
        setError(null);

      } catch (error: unknown) {
        
        const errMessage= getErrorMessage(error);
        
        setIsError(true);
        setError(errMessage);
      } finally {
        setIsLoading(false);
      } 
    };

    execute();   
  }, []);  

  return { data, isLoading, isError, error };    
};

