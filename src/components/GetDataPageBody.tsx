import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';

export interface GetDataPageBodyProps {
    getFileNames: (teamId: string) => {};
    logout: () => {};
}

export default class GetDataPageBody extends React.Component<GetDataPageBodyProps> {

    clickHandler = () => {
        this.props.getFileNames('');
    }
    render() {
        const { logout } = this.props;

        return (
            <div className='ms-welcome__main'>
                <h2 className='ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20'>Select your team to fetch schedule</h2>
                <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.clickHandler}>Get Teams</Button>
                <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={logout}>Sign out from Office 365</Button>
            </div>
        );
    }
}
