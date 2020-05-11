import { AppState } from '../src/components/App';
import { AxiosResponse } from 'axios';

/*
     Interacting with the Office document
*/

export const writeFileNamesToWorksheet = async (result: AxiosResponse,
                                          displayError: (x: string) => void) => {

        return Excel.run( (context: Excel.RequestContext) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            const data = [
                 [result.data.value[0].displayName],
                 [result.data.value[1].displayName],
                 [result.data.value[2].displayName]];

            const range = sheet.getRange('B5:B7');
            range.values = data;
            range.format.autofitColumns();

            return context.sync();
        })
        .catch( (error) => {
            displayError(error.toString());
        });
};

const getMembersMap = (teamOwners: Array<any>, teamMembers: Array<any>) => {
    let membersMap = {};
    teamMembers.forEach((member) => {
        membersMap[member.id] = { ...member, isOwner: false};
    });
    teamOwners.forEach((owner) => {
        membersMap[owner.id] = { ...owner, isOwner: true };
    });
    return membersMap;
};

const getMembersMapByUPN = (membersMap: any) => {
    let membersMapByUPN = {};
    Object.keys(membersMap).forEach((memberId) => {
        const member = membersMap[memberId];
        membersMapByUPN[member.userPrincipalName] = member;
    });
    return membersMapByUPN;
};

const getScheduleGroupNamesMap = (scheduleGroups: Array<any>) => {
    let scheduleGroupNamesMap = {};
    scheduleGroups.forEach((scheduleGroup) => {
        if (scheduleGroup.isActive) {
            scheduleGroupNamesMap[scheduleGroup.displayName] = scheduleGroup;
        }
    });
    return scheduleGroupNamesMap;
};

const isUserInSchedule = (userId: string, scheduleGroupNamesMap: any) => {
    let userInSchedule: boolean = false;
    Object.keys(scheduleGroupNamesMap).forEach((scheduleGroupName) => {
        const scheduleGroup = scheduleGroupNamesMap[scheduleGroupName];
        if (scheduleGroup.userIds.indexOf(userId) >= 0) {
            userInSchedule = true;
        }
    });
    return userInSchedule;
};

const writeGroupMembersToWorksheet = async (membersMap: any,
                                                    scheduleGroupNamesMap: any,
                                                    displayError: (x: string) => void) => {
    return Excel.run( (context: Excel.RequestContext) => {
        context.workbook.worksheets.getItemOrNullObject('Team Membership').delete();
        const sheet = context.workbook.worksheets.add('Team Membership');

        let teamMembersTable = sheet.tables.add('A1:F1', true);
        teamMembersTable.name = 'TeamMembers';
        teamMembersTable.getHeaderRowRange().values = [['User Email', 'DisplayName', 'FirstName', 'LastName', 'isOwner', 'inSchedule']];

        let membersData = [];
        Object.keys(membersMap).forEach((memberId) => {
            const user = membersMap[memberId];
            const userInSchedule = isUserInSchedule(user.id, scheduleGroupNamesMap);
            membersData.push([user.userPrincipalName, user.displayName, user.givenName, user.surname, user.isOwner, userInSchedule]);
        });

        teamMembersTable.rows.add(null, membersData);

        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();

        sheet.activate();
        return context.sync();
    }).catch( (error) => {
        displayError(error.toString());
    });
};

const writeScheduleGroupInformation = async (membersMap: any, scheduleGroupNamesMap: any, displayError: (x: string) => void) => {
    return Excel.run( (context: Excel.RequestContext) => {
        context.workbook.worksheets.getItemOrNullObject('Schedule Membership').delete();
        const sheet = context.workbook.worksheets.add('Schedule Membership');
        let scheduleGroupTable = sheet.tables.add('A1:E1', true);
        scheduleGroupTable.name = 'ScheduleGroups';
        scheduleGroupTable.getHeaderRowRange().values = [['Schedule Group Name', 'User Email', 'Display Name', 'First Name', 'Last Name']];
        // const scheduleGroups = scheduleGroupDataResponse && scheduleGroupDataResponse.data && scheduleGroupDataResponse.data.value || [];
        const scheduleGroupData = [];
        Object.keys(scheduleGroupNamesMap).forEach((scheduleGroupName) => {
            const scheduleGroup = scheduleGroupNamesMap[scheduleGroupName];
            console.log(JSON.stringify(scheduleGroup));
            const membersInGroup = scheduleGroup.userIds || [];
            membersInGroup.forEach((userId) => {
                const user = membersMap[userId];
                console.log(JSON.stringify(user));
                scheduleGroupData.push([scheduleGroup.displayName, user.userPrincipalName, user.displayName, user.givenName, user.surname]);
            });
        });

        scheduleGroupTable.rows.add(null, scheduleGroupData);

        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();

        sheet.activate();
        return context.sync();
    }).catch( (error) => {
        displayError(error.toString());
    });
};


const syncDatatoGraph = async (membersMap: any, scheduleGroupNamesMap: any, displayError: (x: string) => void, pushUpdatedScheduleGroupsToGraph: (scheduleGroupsToSync: any[], selectedTeamId: string) => Promise<any>[], selectedTeamId: string, setState: (x: AppState) => void) => {
    setState({headerMessage: 'Sync data in progress', fileFetch: 'fetchInProcess'});
    return Excel.run( (context: Excel.RequestContext) => {
        // get existing table
        let sheet = context.workbook.worksheets.getItem('Schedule Membership');
        let scheduleGroupsTable = sheet.tables.getItem('ScheduleGroups');
        const scheduleGroupsTableBodyRange = scheduleGroupsTable.getDataBodyRange().load('values');
        const membersMapByUPN = getMembersMapByUPN(membersMap);
        return context.sync().then(async () => {
            for (let i = 0; i < scheduleGroupsTableBodyRange.values.length; i ++) {
                // for each row, check if the row is part of exising schedule
                const scheduleMemberRow = scheduleGroupsTableBodyRange.values[i];
                // TODO: Validate schedule Row
                const scheduleGroupName = scheduleMemberRow[0];
                const userUPN = scheduleMemberRow[1];
                const user = membersMapByUPN[userUPN];
                const scheduleGroup = scheduleGroupNamesMap[scheduleGroupName];

                if (!scheduleGroup) {
                    // no scheduleGroup. Create one
                    let newScheduleGroup = {
                        displayName: scheduleGroupName || "new group",
                        userIds: [user.id],
                        isNew: true,
                        shouldSync: true
                    };
                    scheduleGroupNamesMap[newScheduleGroup.displayName] = newScheduleGroup;
                } else if (scheduleGroup.userIds.indexOf(user.id) < 0) {
                    // check if user is already part of schedule. if not, add the user
                    scheduleGroup.userIds.push(user.id);
                    scheduleGroup.shouldSync = true;
                }
            }
            // filter all schedule groups that need to be synced and update them
            let scheduleGroupsToSync = [];
            Object.keys(scheduleGroupNamesMap).forEach((scheduleGroupName) => {
                const scheduleGroup = scheduleGroupNamesMap[scheduleGroupName];
                if (scheduleGroup.shouldSync) {
                    scheduleGroupsToSync.push(scheduleGroup);
                }
            });
            const updateGroupsPromises = pushUpdatedScheduleGroupsToGraph(scheduleGroupsToSync, selectedTeamId);
            Promise.all(updateGroupsPromises).then(() => {
                setState({headerMessage: 'Sync Complete', fileFetch: 'fetchCompleted'});
                return context.sync();
            });
        });
    }).catch( (error) => {
        displayError(error.toString());
    });
};

export const syncScheduleGroupInfo = async (teamMembersResult : AxiosResponse,
                                            teamOwnersResult: AxiosResponse,
                                            scheduleGroupsResult: AxiosResponse,
                                            displayError: (x: string) => void,
                                            isSync: boolean = false,
                                            pushUpdatedScheduleGroupsToGraph: (scheduleGroupsToSync: any[], selectedTeamId: string) => Promise<any>[],
                                            selectedTeamId: string,
                                            setState: (x: AppState) => void) => {

        const existingTeamMembersArray = teamMembersResult && teamMembersResult.data && teamMembersResult.data.value || [];
        const existingTeamOwnersArray = teamOwnersResult && teamOwnersResult.data && teamOwnersResult.data.value || [];
        const existingScheduleGroupsArray = scheduleGroupsResult && scheduleGroupsResult.data && scheduleGroupsResult.data.value || [];

        const existingMemberMap = getMembersMap(existingTeamOwnersArray, existingTeamMembersArray);
        const existingScheduleGroupNamesMap = getScheduleGroupNamesMap(existingScheduleGroupsArray);

        if (!isSync) {
            // Just get the data and write it to the sheets
            await writeGroupMembersToWorksheet(existingMemberMap, existingScheduleGroupNamesMap, displayError);
            await writeScheduleGroupInformation(existingMemberMap, existingScheduleGroupNamesMap, displayError);
        } else {
            // read the data from the sheets and compare with existing data and write back the delta
            await syncDatatoGraph(existingMemberMap, existingScheduleGroupNamesMap, displayError, pushUpdatedScheduleGroupsToGraph, selectedTeamId, setState);
        }
};

/*
    Managing the dialogs.
*/

let loginDialog: Office.Dialog;
const dialogLoginUrl: string = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/login/login.html';

export const signInO365 = async (setState: (x: AppState) => void,
                                 setToken: (x: string) => void,
                                 displayError: (x: string) => void) => {

    setState({ authStatus: 'loginInProcess' });

    await Office.context.ui.displayDialogAsync(
            dialogLoginUrl,
            {height: 40, width: 30},
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    displayError(`${result.error.code} ${result.error.message}`);
                }
                else {
                    loginDialog = result.value;
                    loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLoginMessage);
                    loginDialog.addEventHandler(Office.EventType.DialogEventReceived, processLoginDialogEvent);
                }
            }
        );

    const processLoginMessage = (arg: {message: string, type: string}) => {

        let messageFromDialog = JSON.parse(arg.message);
        if (messageFromDialog.status === 'success') {

            // We now have a valid access token.
            loginDialog.close();
            setToken(messageFromDialog.result);
            setState( { authStatus: 'loggedIn',
                        headerMessage: 'Select Team' });
        }
        else {
            // Something went wrong with authentication or the authorization of the web application.
            loginDialog.close();
            displayError(messageFromDialog.result);
        }
    };

    const processLoginDialogEvent = (arg) => {
        processDialogEvent(arg, setState, displayError);
    };
};

let logoutDialog: Office.Dialog;
const dialogLogoutUrl: string = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/logout/logout.html';

export const logoutFromO365 = async (setState: (x: AppState) => void,
                                     displayError: (x: string) => void) => {

    Office.context.ui.displayDialogAsync(dialogLogoutUrl,
            {height: 40, width: 30},
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    displayError(`${result.error.code} ${result.error.message}`);
                }
                else {
                    logoutDialog = result.value;
                    logoutDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLogoutMessage);
                    logoutDialog.addEventHandler(Office.EventType.DialogEventReceived, processLogoutDialogEvent);
                }
            }
        );

    const processLogoutMessage = () => {
        logoutDialog.close();
        setState({ authStatus: 'notLoggedIn',
                   headerMessage: 'Welcome' });
    };

    const processLogoutDialogEvent = (arg) => {
        processDialogEvent(arg, setState, displayError);
    };
};

const processDialogEvent = (arg: {error: number, type: string},
                            setState: (x: AppState) => void,
                            displayError: (x: string) => void) => {

    switch (arg.error) {
        case 12002:
            displayError('The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.');
            break;
        case 12003:
            displayError('The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.');
            break;
        case 12006:
            // 12006 means that the user closed the dialog instead of waiting for it to close.
            // It is not known if the user completed the login or logout, so assume the user is
            // logged out and revert to the app's starting state. It does no harm for a user to
            // press the login button again even if the user is logged in.
            setState({ authStatus: 'notLoggedIn',
                       headerMessage: 'Welcome' });
            break;
        default:
            displayError('Unknown error in dialog box.');
            break;
    }
};