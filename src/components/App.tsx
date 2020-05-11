import * as React from 'react';
import { Spinner, SpinnerType } from 'office-ui-fabric-react';
import Header from './Header';
import { HeroListItem } from './HeroList';
import Progress from './Progress';
import StartPageBody from './StartPageBody';
// import GetDataPageBody from './GetDataPageBody';
import SuccessPageBody from './SuccessPageBody';
import OfficeAddinMessageBar from './OfficeAddinMessageBar';
import { getGraphData, updateGraphData, postGraphData } from '../../utilities/microsoft-graph-helpers';
import { writeFileNamesToWorksheet, logoutFromO365, signInO365, syncScheduleGroupInfo } from '../../utilities/office-apis-helpers';


export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
    authStatus?: string;
    fileFetch?: string;
    headerMessage?: string;
    errorMessage?: string;
}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            authStatus: 'notLoggedIn',
            fileFetch: 'notFetched',
            headerMessage: 'Welcome',
            errorMessage: ''
        };

        // Bind the methods that we want to pass to, and call in, a separate
        // module to this component. And rename setState to boundSetState
        // so code that passes boundSetState is more self-documenting.
        this.boundSetState = this.setState.bind(this);
        this.setToken = this.setToken.bind(this);
        this.displayError = this.displayError.bind(this);
        this.login = this.login.bind(this);
    }

    /*
        Properties
    */

    // The access token is not part of state because React is all about the
    // UI and the token is not used to affect the UI in any way.
    accessToken: string;

    listItems: HeroListItem[] = [
        {
            icon: 'PlugConnected',
            primaryText: 'Connects to your teams and schedules'
        },
        {
            icon: 'ExcelDocument',
            primaryText: 'Gets the data from your schedules'
        },
        {
            icon: 'AddNotes',
            primaryText: 'Lets you manage teams and schedules'
        }
    ];

    /*
        Methods
    */

    boundSetState: () => {};

    setToken = (accesstoken: string) => {
        this.accessToken = accesstoken;
    }

    displayError = (error: string) => {
        this.setState({ errorMessage: error });
    }

    // Runs when the user clicks the X to close the message bar where
    // the error appears.
    errorDismissed = () => {
        this.setState({ errorMessage: '' });

        // If the error occured during a "in process" phase (logging in or getting files),
        // the action didn't complete, so return the UI to the preceding state/view.
        this.setState((prevState) => {
            if (prevState.authStatus === 'loginInProcess') {
                return {authStatus: 'notLoggedIn'};
            }
            else if (prevState.fileFetch === 'fetchInProcess') {
                return {fileFetch: 'notFetched'};
            }
            return null;
        });
    }

    login = async () => {
        await signInO365(this.boundSetState, this.setToken, this.displayError);
    }

    logout = async () => {
        await logoutFromO365(this.boundSetState, this.displayError);
    }

    getFileNames = async (_teamId: string) => {
        this.setState({ fileFetch: 'fetchInProcess' });
        getGraphData(

                // Get the `name` property of the first 3 Excel workbooks in the user's OneDrive.
                //"https://graph.microsoft.com/v1.0/me/drive/root/microsoft.graph.search(q = '.xlsx')?$select=name&top=3",
                "https://graph.microsoft.com/v1.0/me/joinedTeams?$select=id,displayName,description",
                this.accessToken
            )
            .then( async (response) => {
                await writeFileNamesToWorksheet(response, this.displayError);
                this.setState({ fileFetch: 'fetched',
                                headerMessage: 'Success' });
            })
            .catch ( (requestError) => {
                // If this runs, then the `then` method did not run, so this error must be
                // from the Axios request in getGraphData, not the Office.js in 
                // writeFileNamesToWorksheet
                this.displayError(requestError);
            });
    }

    pushUpdatedScheduleGroupsToGraph = (scheduleGroupsToSync: any[], selectedTeamId: string) => {
        const updatePromises = scheduleGroupsToSync.map(scheduleGroup => {
            if (scheduleGroup.isNew) {
                const graphUrl =  'https://graph.microsoft.com/v1.0/teams/' + selectedTeamId + '/schedule/schedulingGroups';
                return postGraphData(graphUrl, {...scheduleGroup}, this.accessToken);
            } else {
                const graphUrl = 'https://graph.microsoft.com/v1.0/teams/' + selectedTeamId + '/schedule/schedulingGroups/' + scheduleGroup.id;
                return updateGraphData(graphUrl, { ...scheduleGroup, isActive: true }, this.accessToken);
            }
        });
        return updatePromises;
    }

    syncScheduleInfo = async (teamId: string, isSync: boolean = false) => {
        try {
            // get team members list, team owners list, schedule groups information
            let graphAPICalls = [];
            graphAPICalls.push(getGraphData('https://graph.microsoft.com/v1.0/groups/' + teamId + '/members?$select=id,displayName,mail,userPrincipalName,givenName,surname', this.accessToken));
            graphAPICalls.push(getGraphData('https://graph.microsoft.com/v1.0/groups/' + teamId + '/owners?$select=id,displayName,mail,userPrincipalName,givenName,surname', this.accessToken));
            graphAPICalls.push(getGraphData('https://graph.microsoft.com/v1.0/teams/' + teamId + '/schedule/schedulingGroups', this.accessToken));
            Promise.all(graphAPICalls).then(async (responses) => {
                // await writeGroupMembersToWorksheet(responses[0], responses[1], this.displayError);
                // await writeScheduleGroupInformation(responses[0], responses[2], this.displayError);
                await syncScheduleGroupInfo(responses[0], responses[1], responses[2], this.displayError, isSync, this.pushUpdatedScheduleGroupsToGraph, teamId, this.boundSetState);
            });
        } catch (error) {
            this.displayError(error);
        }
    }

    getTeamsList = async () => {
        let graphResponse = await getGraphData(
            "https://graph.microsoft.com/v1.0/me/joinedTeams?$select=id,displayName,description",
            this.accessToken
        );
        return graphResponse && graphResponse.data && graphResponse.data.value ? graphResponse.data.value : [];
    }

    render() {
        const { title, isOfficeInitialized } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/Shifts_icon.png'
                    message='Please sideload your add-in to see app body.'
                />
            );
        }

        // Set the body of the page based on where the user is in the workflow.
        let body;

        if (this.state.authStatus === 'notLoggedIn') {
            body = ( <StartPageBody login={this.login} listItems={this.listItems}/> );
        }
        else if (this.state.authStatus === 'loginInProcess') {
            body = ( <Spinner className='spinner' type={SpinnerType.large} label='Please sign-in on the pop-up window.' /> );
        }
        else {
            if (this.state.fileFetch === 'notFetched') {
               // body = ( <GetDataPageBody getFileNames={this.getFileNames} logout={this.logout}/> );
               body = (
                    <SuccessPageBody
                        getFileNames={this.getFileNames}
                        logout={this.logout}
                        getTeamsList={this.getTeamsList}
                        syncScheduleInfo={this.syncScheduleInfo}
                    />
                );
            }
            else if (this.state.fileFetch === 'fetchInProcess') {
                body = ( <Spinner className='spinner' type={SpinnerType.large} label='We are syncing the data for you.' /> );
            }
            else {
                body = (
                    <SuccessPageBody
                        getFileNames={this.getFileNames}
                        logout={this.logout}
                        getTeamsList={this.getTeamsList}
                        syncScheduleInfo={this.syncScheduleInfo}
                    />
                );
            }
        }

        return (
            <div>
                { this.state.errorMessage ?
                  (<OfficeAddinMessageBar onDismiss={this.errorDismissed} message={this.state.errorMessage + ' '} />)
                : null }

                <div className='ms-welcome'>
                    <Header logo='assets/Shifts_icon.png' title={this.props.title} message={this.state.headerMessage} />
                    {body}
                </div>
            </div>
        );
    }
}
