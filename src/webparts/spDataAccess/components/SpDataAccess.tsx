import * as React from 'react';

import { ISpDataAccessProps } from './ISpDataAccessProps';

import { ThemeProvider, } from '@fluentui/react';


import { Project } from '../../../interfaces/Project';
import { useQueryProjects } from '../../../services/useQueryProjects';
import { Status } from '../../../interfaces/Status';

const SpDataAccess: React.FC<ISpDataAccessProps> = (props) => {
  
  const {
    theme
  } = props;
  
  const { data, isLoading, error } = useQueryProjects();

  if (isLoading) {
    return <span>Loading...</span>;
  }

  if (error) {
    return <span>Error: {error}</span>;
  }

  return (
    <ThemeProvider theme={theme}>
      <h1>How to get data from SharePoint using pnp V2 and a custom hook</h1>
      <div>
       {data.map((p: Project) => (
          <p key={p.Id}>Id: {p.Id}, Name: {p.Name}, Status: {Status[p.Status]} ({p.Status})</p>
        ))} 
      </div>     
    </ThemeProvider>
  );
};

export default SpDataAccess;
