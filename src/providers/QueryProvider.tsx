/**
 * TanStack Query provider setup for SPFx
 */

import * as React from 'react';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { Logger, LogLevel } from '@pnp/logging';

// Create a client
const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      // Global defaults for all queries
      staleTime: 5 * 60 * 1000, // 5 minutes
      cacheTime: 10 * 60 * 1000, // 10 minutes
      retry: (failureCount, error) => {
        // Don't retry on 4xx errors (client errors)
        if (error && typeof error === 'object' && 'status' in error) {
          const status = (error as any).status;
          if (status >= 400 && status < 500) {
            return false;
          }
        }
        return failureCount < 3;
      },
      refetchOnWindowFocus: false,
      refetchOnReconnect: true,
      onError: (error) => {
        Logger.write(`Query error: ${error}`, LogLevel.Error);
      },
    },
    mutations: {
      // Global defaults for all mutations
      retry: 1,
      onError: (error) => {
        Logger.write(`Mutation error: ${error}`, LogLevel.Error);
      },
    },
  },
});

// Enable devtools in development
if (process.env.NODE_ENV === 'development') {
  // Note: React Query Devtools are not included to keep bundle size small
  // You can add them if needed: npm install @tanstack/react-query-devtools
}

export interface IQueryProviderProps {
  children: React.ReactNode;
}

export const QueryProvider: React.FC<IQueryProviderProps> = ({ children }) => {
  return (
    <QueryClientProvider client={queryClient}>
      {children}
    </QueryClientProvider>
  );
};

// Export the client for direct access if needed
export { queryClient };
