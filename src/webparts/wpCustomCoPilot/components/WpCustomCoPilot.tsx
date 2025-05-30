import * as React from 'react';
import styles from './WpCustomCoPilot.module.scss';
import type { IWpCustomCoPilotProps } from './IWpCustomCoPilotProps';
import { useRef, useEffect } from "react";
import * as ReactWebChat from 'botframework-webchat';
import MSALWrapper from '../SSOAuth/MSALWrapper';
import { Dispatch } from 'redux';
import { Spinner } from '@fluentui/react';


const CoPilotCustomWP: React.FC<IWpCustomCoPilotProps> = (props) => {

  const { botURL, clientID, authority, customScope, userDisplayName, botAvatarImage, botAvatarInitials, userEmail } = props;

  // Check for required properties
  if (!botURL || !clientID || !authority || !customScope) {
    return (
      <section className={styles.wpCustomCoPilot}>
        <div style={{ textAlign: 'center', padding: '1rem', color: 'red' }}>
          Please configure webpart properties
        </div>
      </section>
    );
  }

  // constructing URL using regional settings
  const environmentEndPoint = botURL.slice(0, botURL.indexOf('/powervirtualagents'));
  const apiVersion = botURL.slice(botURL.indexOf('api-version')).split('=')[1];
  const regionalChannelSettingsURL = `${environmentEndPoint}/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`;

  // Using refs instead of IDs to get the webchat and loading spinner elements
  const webChatRef = useRef<HTMLDivElement>(null);
  const loadingSpinnerRef = useRef<HTMLDivElement>(null);

  // A utility function that extracts the OAuthCard resource URI from the incoming activity or return undefined
  function getOAuthCardResourceUri(activity: any): string | undefined {
    const attachment = activity?.attachments?.[0];
    if (attachment?.contentType === 'application/vnd.microsoft.card.oauth' && attachment.content.tokenExchangeResource) {
      return attachment.content.tokenExchangeResource.uri;
    }
  }

  const onDidMount = async () => {
        
    const MSALWrapperInstance = new MSALWrapper(props.clientID, props.authority);

    // Trying to get token if user is already signed-in
    let responseToken = await MSALWrapperInstance.handleLoggedInUser([props.customScope], props.userEmail);

    if (!responseToken) {
        // Trying to get token if user is not signed-in
        responseToken = await MSALWrapperInstance.acquireAccessToken([props.customScope], props.userEmail);
    }

    const token = responseToken?.accessToken || null;

    // Get the regional channel URL
    let regionalChannelURL;

    const regionalResponse = await fetch(regionalChannelSettingsURL);
    if(regionalResponse.ok){
        const data = await regionalResponse.json();
        regionalChannelURL = data.channelUrlsById.directline;
    }
    else {
        console.error(`HTTP error! Status: ${regionalResponse.status}`);
    }


    // Create DirectLine object
    let directline: any;

    const response = await fetch(botURL);
    
    if (response.ok) {
        const conversationInfo = await response.json();
        directline = ReactWebChat.createDirectLine({
        token: conversationInfo.token, 
        domain: regionalChannelURL + 'v3/directline'
    });
    } else {
    console.error(`HTTP error! Status: ${response.status}`);
    }

    const store = ReactWebChat.createStore(
        {},
           ({ dispatch }: { dispatch: Dispatch }) => (next: any) => (action: any) => {
               
            // Checking whether we should greet the user
            if (props.greet)
            {
                if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
                    console.log("Action:" + action.type); 
                        dispatch({
                            meta: {
                                method: "keyboard",
                              },
                                payload: {
                                  activity: {
                                          channelData: {
                                              postBack: true,
                                          },
                                          //Web Chat will show the 'Greeting' System Topic message which has a trigger-phrase 'hello'
                                          name: 'startConversation',
                                          type: "event"
                                      },
                              },
                              type: "DIRECT_LINE/POST_ACTIVITY",
                          });
                          return next(action);
                      }
                }
                
                // Checking whether the bot is asking for authentication
                if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
                    const activity = action.payload.activity;
                    if (activity.from && activity.from.role === 'bot' &&
                    (getOAuthCardResourceUri(activity))){
                      directline.postActivity({
                        type: 'invoke',
                        name: 'signin/tokenExchange',
                        value: {
                          id: activity.attachments[0].content.tokenExchangeResource.id,
                          connectionName: activity.attachments[0].content.connectionName,
                          token
                        },
                        "from": {
                          id: props.userEmail,
                          name: props.userFriendlyName,
                          role: "user"
                        }
                            }).subscribe(
                                (id: any) => {
                                  if(id === "retry"){
                                    // bot was not able to handle the invoke, so display the oauthCard (manual authentication)
                                    console.log("bot was not able to handle the invoke, so display the oauthCard")
                                        return next(action);
                                  }
                                },
                                (error: any) => {
                                  // an error occurred to display the oauthCard (manual authentication)
                                  console.log("An error occurred so display the oauthCard");
                                      return next(action);
                                }
                              )
                            // token exchange was successful, do not show OAuthCard
                            return;
                    }
                  } else {
                    return next(action);
                  }
                
                return next(action);
            }
        );

        const avatarOptions = botAvatarImage && botAvatarInitials ? {
          botAvatarImage: botAvatarImage,
          botAvatarInitials: botAvatarInitials,
          userAvatarImage: `/_layouts/15/userphoto.aspx?size=S&username=${userEmail}`,
          userAvatarInitials: userDisplayName.charAt(0)
        } : {};
        // hide the upload button - other style options can be added here
        const canvasStyleOptions = {
            hideUploadButton: true,
            rootHeight: '100%',
            rootWidth: '100%',
            botAvatarBackgroundColor: '#fff',
            userAvatarBackgroundColor: '#fff',
            bubbleBackground: '#EBEBED',
            bubbleTextColor: '#000',
            bubbleFromUserBackground: '#0057B8',
            bubbleFromUserTextColor: '#fff',
            sendBoxBackground: '#F3F4F6',
            ...avatarOptions
        }
    
        // Render webchat
        if (token && directline) {
            if (webChatRef.current && loadingSpinnerRef.current) {
                webChatRef.current.style.minHeight = '40vh';
                loadingSpinnerRef.current.style.display = 'none';
                ReactWebChat.renderWebChat(
                    {
                        directLine: directline,
                        store: store,
                        username: userDisplayName,
                        styleOptions: canvasStyleOptions,
                        userID: props.userEmail,
                        sendTypingIndicator: true,
                    },
                webChatRef.current
                );
            } else {
                console.error("Webchat or loading spinner not found");
            }
    }

};
  useEffect(() => {
    console.log('Component mounted');
    onDidMount();
    // Cleanup function (optional, similar to componentWillUnmount)
    return () => {
      console.log('Component unmounted');
    };
  }, []); // Empty dependency array ensures this runs only once

  return (
    <section className={`${styles.wpCustomCoPilot}`} style={{ width: props.width ? `${props.width}px` : '100%' }}>
      <div
        className={styles.header}
        style={{
          background: props.headerBgColor || '#009FDB',
          color: props.headerTextColor || '#fff',
          padding: '1rem',
          paddingLeft: props.headerPaddingLeft ? `${props.headerPaddingLeft}px` : '0',
          borderRadius: '4px 4px 0 0',
          fontWeight: 'bold',
          fontSize: props.headerFontSize ? `${props.headerFontSize}px` : '1.3rem',
          letterSpacing: '0.5px',
          height: props.headerHeight ? `${props.headerHeight}px` : undefined,
          minHeight: props.headerHeight ? `${props.headerHeight}px` : undefined,
          maxHeight: props.headerHeight ? `${props.headerHeight}px` : undefined
        }}
      >
        FiTo AI (Powered by Ask AT&T)
      </div>
      <div
        className={styles.chatContainer}
        id="chatContainer"
        style={{
          paddingTop: props.chatContainerPaddingTop ? `${props.chatContainerPaddingTop}px` : '0',
          height: props.height ? `${props.height}px` : '400px',
          width: props.width ? `${props.width}px` : '100%'
        }}
      >
        <div ref={webChatRef} role="main" className={styles.webChat}></div>
        <div ref={loadingSpinnerRef}><Spinner label="Loading..." style={{ paddingTop: "1rem", paddingBottom: "1rem" }} /></div>
      </div>
    </section>
  );
}

export default class WpCustomCoPilot extends React.Component<IWpCustomCoPilotProps, {}> {

  constructor(props: IWpCustomCoPilotProps) {
    super(props);
    console.log(props);
  }

  public render(): React.ReactElement<IWpCustomCoPilotProps> {
    return (
      <CoPilotCustomWP {...this.props} />
    );
  }
}


 