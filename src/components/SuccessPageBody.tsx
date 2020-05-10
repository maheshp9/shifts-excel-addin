import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface SuccessPageBodyProps {
    getFileNames: (teamId: string) => {};
    logout: () => {};
    getTeamsList: () => Promise<any>;
    addScheduleInfo?: (teamId: string) => {};
}

export interface SuccessPagebodyAppSate {
    dataFetchStatus?: boolean;
    selectedTeamId?: string;
}

export default class SuccessPageBody extends React.Component<SuccessPageBodyProps, SuccessPagebodyAppSate> {
    private teamsList: Array<any> = null;
    constructor(props, context) {
        super(props, context);
        this.state = {
            dataFetchStatus: false,
            selectedTeamId: null
        };
        this.teamsList = [];
    }

    async componentDidMount() {
        this.teamsList = await this.props.getTeamsList();
        console.log(JSON.stringify(this.teamsList));
        this.setState({dataFetchStatus: true});
    }

    getDropDownOptions = (teamsList : Array<any>) => {
        let dropDownOptions = [];
        teamsList.forEach((teamItem: any) => dropDownOptions.push({
            'key': teamItem.id,
            'text': teamItem.displayName
        }));
        return dropDownOptions;
    }

    onTeamSelected = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        if (option) {
            // this.props.addScheduleInfo(option.key as string);
            this.setState({selectedTeamId: option.key as string });
            console.log(option.key);
        }
    }

    getScheduleButtonHandler = () => {
        this.props.addScheduleInfo(this.state.selectedTeamId);
    }

    renderBody = () => {
        const { logout } = this.props;
        if (this.state.dataFetchStatus) {
            return (
                <>
                    <Dropdown placeHolder='Select a team' options={ this.getDropDownOptions(this.teamsList || []) }  onChange={this.onTeamSelected } />
                    <h2 className='ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20'>Selected Team: { this.state.selectedTeamId || '' } </h2>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.getScheduleButtonHandler}>Get Schedules</Button>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={logout}>Sign out from Office 365</Button>
                </>
            );
        } else {
            return (
                <Spinner label='Fetching teams data'/>
            );
        }
    }


    render() {

        return (
            <div className='ms-welcome__main'>
                { this.renderBody() }
            </div>
        );
    }
}
