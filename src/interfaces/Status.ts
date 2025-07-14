/**
 * Enhanced Status enum with additional status types and utilities
 */

export enum Status {
  NotStarted = 0,
  InProgress = 1,
  Completed = 2,
  OnHold = 3,
  Cancelled = 4,
  UnderReview = 5,
  Approved = 6,
  Rejected = 7,
}

export interface StatusInfo {
  value: Status;
  label: string;
  color: string;
  backgroundColor: string;
  description: string;
}

export const StatusConfig: Record<Status, StatusInfo> = {
  [Status.NotStarted]: {
    value: Status.NotStarted,
    label: 'Not Started',
    color: '#605e5c',
    backgroundColor: '#f3f2f1',
    description: 'Project has not been started yet',
  },
  [Status.InProgress]: {
    value: Status.InProgress,
    label: 'In Progress',
    color: '#0078d4',
    backgroundColor: '#deecf9',
    description: 'Project is currently being worked on',
  },
  [Status.Completed]: {
    value: Status.Completed,
    label: 'Completed',
    color: '#107c10',
    backgroundColor: '#dff6dd',
    description: 'Project has been successfully completed',
  },
  [Status.OnHold]: {
    value: Status.OnHold,
    label: 'On Hold',
    color: '#d13438',
    backgroundColor: '#fde7e9',
    description: 'Project is temporarily paused',
  },
  [Status.Cancelled]: {
    value: Status.Cancelled,
    label: 'Cancelled',
    color: '#a4262c',
    backgroundColor: '#fde7e9',
    description: 'Project has been cancelled',
  },
  [Status.UnderReview]: {
    value: Status.UnderReview,
    label: 'Under Review',
    color: '#8764b8',
    backgroundColor: '#f4f1fa',
    description: 'Project is being reviewed',
  },
  [Status.Approved]: {
    value: Status.Approved,
    label: 'Approved',
    color: '#498205',
    backgroundColor: '#e9f7df',
    description: 'Project has been approved',
  },
  [Status.Rejected]: {
    value: Status.Rejected,
    label: 'Rejected',
    color: '#d13438',
    backgroundColor: '#fde7e9',
    description: 'Project has been rejected',
  },
};

/**
 * Gets the status configuration for a given status value
 * @param status - The status value
 * @returns StatusInfo object with display properties
 */
export const getStatusInfo = (status: Status): StatusInfo => {
  return StatusConfig[status] || StatusConfig[Status.NotStarted];
};

/**
 * Gets all available status options for dropdowns/selects
 * @returns Array of status options
 */
export const getStatusOptions = () => {
  return Object.values(StatusConfig).map(config => ({
    key: config.value.toString(),
    text: config.label,
    data: config,
  }));
};

/**
 * Checks if a status represents an active/working state
 * @param status - The status to check
 * @returns True if the status is active
 */
export const isActiveStatus = (status: Status): boolean => {
  return status === Status.InProgress || status === Status.UnderReview;
};

/**
 * Checks if a status represents a completed state
 * @param status - The status to check
 * @returns True if the status is completed
 */
export const isCompletedStatus = (status: Status): boolean => {
  return status === Status.Completed || status === Status.Approved;
};

/**
 * Checks if a status represents a blocked/stopped state
 * @param status - The status to check
 * @returns True if the status is blocked
 */
export const isBlockedStatus = (status: Status): boolean => {
  return status === Status.OnHold || status === Status.Cancelled || status === Status.Rejected;
};

/**
 * Gets the next logical status transitions for a given status
 * @param currentStatus - The current status
 * @returns Array of possible next statuses
 */
export const getNextStatusOptions = (currentStatus: Status): Status[] => {
  switch (currentStatus) {
    case Status.NotStarted:
      return [Status.InProgress, Status.OnHold, Status.Cancelled];
    case Status.InProgress:
      return [Status.Completed, Status.OnHold, Status.UnderReview];
    case Status.OnHold:
      return [Status.InProgress, Status.Cancelled];
    case Status.UnderReview:
      return [Status.Approved, Status.Rejected, Status.InProgress];
    case Status.Approved:
      return [Status.InProgress];
    case Status.Rejected:
      return [Status.InProgress, Status.Cancelled];
    case Status.Completed:
      return [Status.UnderReview]; // For re-opening if needed
    case Status.Cancelled:
      return [Status.NotStarted]; // For restarting
    default:
      return [];
  }
};
