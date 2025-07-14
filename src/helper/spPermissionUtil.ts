// import { sp } from '@pnp/sp';
// import '@pnp/sp/webs';
// import '@pnp/sp/lists';
// import '@pnp/sp/items';
// import '@pnp/sp/security';
// import { IRoleAssignment, IRoleAssignmentInfo, IRoleAssignments, IRoleDefinition, IRoleDefinitionInfo, IRoleDefinitions } from '@pnp/sp/security';


// export async function IsUserInRole(roleName: string): Promise<boolean> {
    
//     const roleDefInfo: IRoleDefinitionInfo = await sp.web.roleDefinitions.getByName(roleName).get();

//     console.debug("roledef: " + roleDefInfo.Id + " " + roleDefInfo.Name);

//     const perms2 = await sp.web.getCurrentUserEffectivePermissions();

//     console.debug(perms2);

//     const ras: IRoleAssignments = await sp.web.roleAssignments.get();

//      console.debug(ras);   

//     return true;
// }
