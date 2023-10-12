import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, authentication } from "@microsoft/teams-js";
import Axios from "axios";
import { ICustomer } from "../../model/ICustomer";

/**
 * Implementation of the Meeting Data content page
 */
export const MeetingDataTab = () => {

  const [{ inTeams, theme, context }] = useTeams();
  const [entityId, setEntityId] = useState<string | undefined>();
  const [customer, setCustomer] = useState<ICustomer>();
  const [error, setError] = useState<string>();

  const loadCustomer = (idToken: string, meetingID: string) => {
    Axios.get(`https://${process.env.PUBLIC_HOSTNAME}/api/customer/${meetingID}`, {
                responseType: "json",
                headers: {
                  Authorization: `Bearer ${idToken}`
                }
    }).then(result => {
      if (result.data) {
        setCustomer(result.data);
      }     
    })
    .catch((error) => {
      console.log(error);
    })
  };

  useEffect(() => {
    if (inTeams === true) {
      authentication.getAuthToken({
        resources: [`api://${process.env.PUBLIC_HOSTNAME}/${process.env.TAB_APP_ID}`],
        silent: false
      } as authentication.AuthTokenRequestParameters).then(token => {
        const meetingID: string = '19:meeting_NWM3OTY5OTItOGY2NS00YzQ0LTlkZjgtZjMwNDc4NjUwMTAw@thread.v2';
        // const meetingID: string = context?.meeting?.id ? context?.meeting?.id : '';
        loadCustomer(token, meetingID);
        app.notifySuccess();
      }).catch(message => {
        setError(message);
        app.notifyFailure({
          reason: app.FailedReason.AuthFailed,
          message
        });
      });
    } else {
      setEntityId("Not in Microsoft Teams");
    }
  }, [inTeams]);

  /**
   * The render() method to create the UI of the tab
   */
  return (
    <Provider theme={theme}>
      <Flex fill={true} column styles={{
        padding: ".8rem 0 .8rem .5rem"
      }}>
        <Flex.Item>
          <Header content="Customer Info" />
        </Flex.Item>
        <Flex.Item>                                      
          <div className="gridTable">      
            <div className="gridRow">
              <div className="gridCell3">
                <label>Name</label>
              </div>      
              <div className="gridCell9">
                <label id="customerName" className="infoData">{customer && customer.Name ? customer.Name : ''}</label>
              </div>
            </div>
            <div className="gridRow">
              <div className="gridCell3">
                <label>Phone</label>
              </div>
              <div className="gridCell9">
                <label id="customerPhone" className="infoData">{customer && customer.Phone ? customer.Phone : ''}</label>
              </div>
            </div>
            <div className="gridRow">
              <div className="gridCell3">
                <label>Email</label>
              </div>
              <div className="gridCell9">
                <label id="customerEmail" className="infoData">{customer && customer.Email ? customer.Email : ''}</label>
              </div>
            </div>
            <div className="gridRow">
              <div className="gridCell3">
                <label>ID</label>
              </div>
              <div className="gridCell9">
                <label id="customerID" className="infoData">{customer && customer.Id ? customer.Id : ''}</label>
              </div>
            </div>
          </div>
        </Flex.Item>        
      </Flex>
    </Provider>
  );
};
