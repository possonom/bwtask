/**
 * SharePoint permissions helper with enhanced functionality
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/security';
import '@pnp/sp/site-users';
import { Logger, LogLevel } from "@pnp/logging";
import { getErrorMessage, logError } from './errorHelper';

export enum SPPermissionLevel {
  FullControl = 'Full Control',
  Design = 'Design',
  Edit = 'Edit',
  Contribute = 'Contribute',
  Read = 'Read',
  ViewOnly = 'View Only',
  RestrictedRead = 'Restricted Read',
}

export interface IUserPermissions {
  userId: number;
  loginName: string;
  email: string;
  displayName: string;
  permissions: string[];
  hasFullControl: boolean;
  canEdit: boolean;
  canRead: boolean;
}

export interface IListPermissions {
  listId: string;
  listTitle: string;
  userPermissions: IUserPermissions;
  inheritedFromParent: boolean;
}

/**
 * Checks if the current user has specific permissions on a list
 * @param listTitle - The title of the SharePoint list
 * @param permission - The permission level to check
 * @returns Promise<boolean> - True if user has the permission
 */
export const hasListPermission = async (listTitle: string, permission: SPPermissionLevel): Promise<boolean> => {
  try {
    Logger.write(`Checking ${permission} permission for list: ${listTitle}`, LogLevel.Info);
    
    const list = sp.web.lists.getByTitle(listTitle);
    const currentUser = await sp.web.currentUser.get();
    
    // Get user's effective permissions on the list
    const userPermissions = await list.getUserEffectivePermissions(currentUser.LoginName);
    
    // Check specific permission based on the level
    let hasPermission = false;
    
    switch (permission) {
      case SPPermissionLevel.FullControl:
        hasPermission = sp.web.hasPermissions(userPermissions, 65);
        break;
      case SPPermissionLevel.Edit:
      case SPPermissionLevel.Contribute:
        hasPermission = sp.web.hasPermissions(userPermissions, 30);
        break;
      case SPPermissionLevel.Read:
        hasPermission = sp.web.hasPermissions(userPermissions, 1);
        break;
      default:
        hasPermission = sp.web.hasPermissions(userPermissions, 1);
    }
    
    Logger.write(`User ${currentUser.LoginName} ${hasPermission ? 'has' : 'does not have'} ${permission} permission on ${listTitle}`, LogLevel.Info);
    return hasPermission;
    
  } catch (error) {
    logError(error, `hasListPermission for ${listTitle}`);
    return false;
  }
};

/**
 * Gets detailed permission information for the current user on a specific list
 * @param listTitle - The title of the SharePoint list
 * @returns Promise<IListPermissions> - Detailed permission information
 */
export const getCurrentUserListPermissions = async (listTitle: string): Promise<IListPermissions | null> => {
  try {
    Logger.write(`Getting current user permissions for list: ${listTitle}`, LogLevel.Info);
    
    const list = sp.web.lists.getByTitle(listTitle);
    const currentUser = await sp.web.currentUser.get();
    
    // Get user's effective permissions
    const userPermissions = await list.getUserEffectivePermissions(currentUser.LoginName);
    
    // Check various permission levels
    const permissions: string[] = [];
    const hasFullControl = sp.web.hasPermissions(userPermissions, 65);
    const canEdit = sp.web.hasPermissions(userPermissions, 30);
    const canRead = sp.web.hasPermissions(userPermissions, 1);
    
    if (hasFullControl) permissions.push(SPPermissionLevel.FullControl);
    if (canEdit) permissions.push(SPPermissionLevel.Edit);
    if (canRead) permissions.push(SPPermissionLevel.Read);
    
    // Check if permissions are inherited
    const listInfo = await list.select('HasUniqueRoleAssignments').get();
    
    const result: IListPermissions = {
      listId: listInfo.Id,
      listTitle: listTitle,
      userPermissions: {
        userId: currentUser.Id,
        loginName: currentUser.LoginName,
        email: currentUser.Email,
        displayName: currentUser.Title,
        permissions,
        hasFullControl,
        canEdit,
        canRead,
      },
      inheritedFromParent: !listInfo.HasUniqueRoleAssignments,
    };
    
    Logger.write(`Successfully retrieved permissions for ${currentUser.LoginName} on ${listTitle}`, LogLevel.Info);
    return result;
    
  } catch (error) {
    logError(error, `getCurrentUserListPermissions for ${listTitle}`);
    return null;
  }
};

/**
 * Checks if the current user can perform CRUD operations on a list
 * @param listTitle - The title of the SharePoint list
 * @returns Promise with CRUD permissions
 */
export const getCRUDPermissions = async (listTitle: string) => {
  try {
    Logger.write(`Checking CRUD permissions for list: ${listTitle}`, LogLevel.Info);
    
    const [canRead, canCreate, canEdit, canDelete] = await Promise.all([
      hasListPermission(listTitle, SPPermissionLevel.Read),
      hasListPermission(listTitle, SPPermissionLevel.Contribute),
      hasListPermission(listTitle, SPPermissionLevel.Edit),
      hasListPermission(listTitle, SPPermissionLevel.FullControl),
    ]);
    
    return {
      canRead,
      canCreate,
      canEdit,
      canDelete,
      canManage: await hasListPermission(listTitle, SPPermissionLevel.FullControl),
    };
    
  } catch (error) {
    logError(error, `getCRUDPermissions for ${listTitle}`);
    return {
      canRead: false,
      canCreate: false,
      canEdit: false,
      canDelete: false,
      canManage: false,
    };
  }
};

/**
 * Checks if a list exists and the user has access to it
 * @param listTitle - The title of the SharePoint list
 * @returns Promise<boolean> - True if list exists and user has access
 */
export const canAccessList = async (listTitle: string): Promise<boolean> => {
  try {
    Logger.write(`Checking access to list: ${listTitle}`, LogLevel.Info);
    
    const list = await sp.web.lists.getByTitle(listTitle).get();
    
    if (list) {
      Logger.write(`List ${listTitle} exists and is accessible`, LogLevel.Info);
      return true;
    }
    
    return false;
    
  } catch (error) {
    const errorMessage = getErrorMessage(error);
    
    if (errorMessage.includes('does not exist') || errorMessage.includes('Not Found')) {
      Logger.write(`List ${listTitle} does not exist`, LogLevel.Warning);
    } else if (errorMessage.includes('Access denied') || errorMessage.includes('Unauthorized')) {
      Logger.write(`Access denied to list ${listTitle}`, LogLevel.Warning);
    } else {
      logError(error, `canAccessList for ${listTitle}`);
    }
    
    return false;
  }
};

/**
 * Gets information about the current user's site permissions
 * @returns Promise with site-level permission information
 */
export const getCurrentUserSitePermissions = async () => {
  try {
    Logger.write('Getting current user site permissions', LogLevel.Info);
    
    const currentUser = await sp.web.currentUser.get();
    const userPermissions = await sp.web.getUserEffectivePermissions(currentUser.LoginName);
    
    const isSiteAdmin = sp.web.hasPermissions(userPermissions, 65);
    const canManageLists = sp.web.hasPermissions(userPermissions, 32);
    const canCreateLists = sp.web.hasPermissions(userPermissions, 16);
    const canBrowse = sp.web.hasPermissions(userPermissions, 1);
    
    return {
      user: {
        id: currentUser.Id,
        loginName: currentUser.LoginName,
        email: currentUser.Email,
        displayName: currentUser.Title,
      },
      permissions: {
        isSiteAdmin,
        canManageLists,
        canCreateLists,
        canBrowse,
      },
    };
    
  } catch (error) {
    logError(error, 'getCurrentUserSitePermissions');
    return null;
  }
};

/**
 * Utility function to format permission information for display
 * @param permissions - The permission object to format
 * @returns Formatted permission string
 */
export const formatPermissions = (permissions: IUserPermissions): string => {
  if (permissions.hasFullControl) {
    return 'Full Control';
  } else if (permissions.canEdit) {
    return 'Edit';
  } else if (permissions.canRead) {
    return 'Read Only';
  } else {
    return 'No Access';
  }
};
