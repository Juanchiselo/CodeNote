import * as React from 'react';
// import { Button, ButtonType } from 'office-ui-fabric-react';
import Splash from './Splash';
// import HeroList, { HeroListItem } from './HeroList';
import Progress from './Progress';

import * as OfficeHelpers from '@microsoft/office-js-helpers';

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
    // listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            listItems: []
        };
    }

    componentDidMount() {


        this.setState({
            listItems: [
                {
                    icon: 'Ribbon',
                    primaryText: 'Achieve more with Office integration'
                },
                {
                    icon: 'Unlock',
                    primaryText: 'Unlock features and functionality'
                },
                {
                    icon: 'Design',
                    primaryText: 'Create and visualize like a pro'
                }
            ]
        });
    }

    click = async () => {
        try {
            await OneNote.run(async context => {
                /**
                 * Insert your OneNote code here
                 */
                return context.sync();
            });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }

    render() {
        const {
            title,
            isOfficeInitialized,
        } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/codenote-filled.png'
                    message='Please sideload your addin to see app body.'
                />
            );
        }

        return (
            <div className='splash-screen'>
                <Splash logo='assets/codenote-filled.png' 
                        title={this.props.title} 
                        message='CodeNote' />
                {/* <HeroList message='Discover what CodeNote can do for you today!' 
                          items={this.state.listItems}>
                    <p className='ms-font-l'>Modify the source files, then click <b>Run</b>.</p>
                    <Button className='ms-welcome__action' 
                            buttonType={ButtonType.hero}
                            iconProps={{ iconName: 'ChevronRight' }}
                            onClick={this.click}>Run</Button>
                </HeroList> */}
            </div>
        );
    }
}
