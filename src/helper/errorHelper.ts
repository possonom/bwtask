/**
 * Enhanced error helper with better error handling and logging
 */

import { Logger, LogLevel } from "@pnp/logging";

export interface IErrorInfo {
  message: string;
  code?: string;
  details?: any;
  timestamp: Date;
}

/**
 * Extracts a user-friendly error message from various error types
 * @param error - The error object to process
 * @returns A formatted error message string
 */
export const getErrorMessage = (error: unknown): string => {
  if (!error) {
    return 'An unknown error occurred';
  }

  // Handle string errors
  if (typeof error === 'string') {
    return error;
  }

  // Handle Error objects
  if (error instanceof Error) {
    return error.message || 'An error occurred';
  }

  // Handle SharePoint/PnP specific errors
  if (typeof error === 'object' && error !== null) {
    const errorObj = error as any;
    
    // Check for SharePoint error structure
    if (errorObj.data && errorObj.data.responseBody) {
      try {
        const responseBody = JSON.parse(errorObj.data.responseBody);
        if (responseBody.error && responseBody.error.message) {
          return responseBody.error.message.value || responseBody.error.message;
        }
      } catch (parseError) {
        Logger.write(`Error parsing SharePoint error response: ${parseError}`, LogLevel.Warning);
      }
    }

    // Check for common error properties
    if (errorObj.message) {
      return errorObj.message;
    }

    if (errorObj.description) {
      return errorObj.description;
    }

    if (errorObj.statusText) {
      return errorObj.statusText;
    }

    // Handle HTTP errors
    if (errorObj.status) {
      return `HTTP Error ${errorObj.status}: ${errorObj.statusText || 'Request failed'}`;
    }
  }

  // Fallback: stringify the error
  try {
    return JSON.stringify(error);
  } catch (stringifyError) {
    return 'An unknown error occurred (unable to serialize error details)';
  }
};

/**
 * Creates a detailed error info object for logging and debugging
 * @param error - The error to process
 * @param context - Additional context about where the error occurred
 * @returns An IErrorInfo object with detailed error information
 */
export const createErrorInfo = (error: unknown, context?: string): IErrorInfo => {
  const errorInfo: IErrorInfo = {
    message: getErrorMessage(error),
    timestamp: new Date(),
  };

  if (typeof error === 'object' && error !== null) {
    const errorObj = error as any;
    
    if (errorObj.code) {
      errorInfo.code = errorObj.code;
    }
    
    if (errorObj.status) {
      errorInfo.code = errorObj.status.toString();
    }

    // Include additional details for debugging
    errorInfo.details = {
      context,
      originalError: errorObj,
      stack: errorObj.stack,
      name: errorObj.name,
    };
  }

  return errorInfo;
};

/**
 * Logs an error with appropriate level and context
 * @param error - The error to log
 * @param context - Context about where the error occurred
 * @param level - Log level (defaults to Error)
 */
export const logError = (error: unknown, context?: string, level: LogLevel = LogLevel.Error): void => {
  const errorInfo = createErrorInfo(error, context);
  const logMessage = context 
    ? `${context}: ${errorInfo.message}`
    : errorInfo.message;

  Logger.write(logMessage, level);
  
  // Log additional details in verbose mode
  if (errorInfo.details) {
    Logger.write(`Error details: ${JSON.stringify(errorInfo.details, null, 2)}`, LogLevel.Verbose);
  }
};

/**
 * Handles async operations with error logging
 * @param operation - The async operation to execute
 * @param context - Context for error logging
 * @returns Promise that resolves to the operation result or throws with enhanced error info
 */
export const handleAsyncOperation = async <T>(
  operation: () => Promise<T>,
  context: string
): Promise<T> => {
  try {
    Logger.write(`${context} - Starting operation`, LogLevel.Info);
    const result = await operation();
    Logger.write(`${context} - Operation completed successfully`, LogLevel.Info);
    return result;
  } catch (error) {
    logError(error, context);
    throw error;
  }
};

/**
 * Creates a user-friendly error message for display in UI
 * @param error - The error to format
 * @param fallbackMessage - Fallback message if error cannot be processed
 * @returns A user-friendly error message
 */
export const getUserFriendlyErrorMessage = (error: unknown, fallbackMessage = 'Something went wrong. Please try again.'): string => {
  const message = getErrorMessage(error);
  
  // Map technical errors to user-friendly messages
  const errorMappings: { [key: string]: string } = {
    'Network request failed': 'Unable to connect to the server. Please check your internet connection.',
    'Unauthorized': 'You do not have permission to perform this action.',
    'Forbidden': 'Access denied. Please contact your administrator.',
    'Not Found': 'The requested resource was not found.',
    'Internal Server Error': 'A server error occurred. Please try again later.',
  };

  // Check if we have a user-friendly mapping
  for (const [technical, friendly] of Object.entries(errorMappings)) {
    if (message.includes(technical)) {
      return friendly;
    }
  }

  // Return the original message if it's already user-friendly, otherwise use fallback
  return message.length > 100 ? fallbackMessage : message;
};
