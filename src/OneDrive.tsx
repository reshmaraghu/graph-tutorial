import React from 'react';
import { Table } from 'reactstrap';
import moment from 'moment';
import { DriveItem } from 'microsoft-graph';
import { config } from './Config';
import { getDriveItems } from './GraphService';
import withAuthProvider, { AuthComponentProps } from './AuthProvider';

interface OneDriveState {
  driveItems: DriveItem[];
}

// Helper function to format Graph date/time
function formatDateTime(dateTime: string | undefined) {
  if (dateTime !== undefined) {
    return moment.utc(dateTime).local().format('M/D/YY h:mm A');
  }
}

class OneDrive extends React.Component<AuthComponentProps, OneDriveState> {
  constructor(props: any) {
    super(props);

    this.state = {
      driveItems: []
    };
  }

  async componentDidMount() {
    try {
      // Get the user's access token
      var accessToken = await this.props.getAccessToken(config.scopes);
      // Get the items in user's OneDrive
      var driveItems = await getDriveItems(accessToken);
      // Update the array of events in state
      this.setState({driveItems: driveItems.value});
    }
    catch(err) {
      this.props.setError('ERROR', JSON.stringify(err));
    }
  }

  // <renderSnippet>
  render() {
    return (
      <div>
        <h1>OneDrive</h1>
        <Table>
          <thead>
            <tr>
              <th scope="col">Name</th>
              <th scope="col">Created Date Time</th>
            </tr>
          </thead>
          <tbody>
            {this.state.driveItems.map(
              function(driveItem: DriveItem){
                return(
                  <tr key={driveItem.id}>
                    <td>{driveItem.name}</td>
                    <td>{formatDateTime(driveItem.createdDateTime)}</td>
                  </tr>
                );
              })}
          </tbody>
        </Table>
      </div>
    );
  }
  // </renderSnippet>
}

export default withAuthProvider(OneDrive);
