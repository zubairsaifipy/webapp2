import * as React from "react";
import { useBoolean, useId } from "@uifabric/react-hooks";
import * as ReactWebChat from "botframework-webchat";
import { Dialog, DialogType } from "office-ui-fabric-react/lib/Dialog";
// import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { Dispatch } from "redux";
import { useRef } from "react";

import { IChatbotProps } from "./IChatBotProps";
import MSALWrapper from "./MSALWrapper";
import "./Chatbot.css";
import { Image } from "office-ui-fabric-react";

export const PVAChatbotDialog: React.FunctionComponent<IChatbotProps> = (
  props
) => {
  // Dialog properties and states
  const dialogContentProps = {
    type: DialogType.normal,
    title: props.botName,
    closeButtonAriaLabel: "Close",
  };

  const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);
  const labelId: string = useId("dialogLabel");
  const subTextId: string = useId("subTextLabel");

  const modalProps = React.useMemo(
    () => ({
      isBlocking: false,
    }),
    [labelId, subTextId]
  );

  const fixedButtonIcon = require("../images/msgIcon.png");
  const botIcon = require("../images/bot.jpg");
  const profileIcon = require("../images/profile.jpg");
  const closeIcon = require("../images/closeIcon.svg");
  const chatLoader = require("../images/loader.svg");

  // Your bot's token endpoint
  const botURL = props.botURL;

  // constructing URL using regional settings
  const environmentEndPoint = botURL.slice(
    0,
    botURL.indexOf("/powervirtualagents")
  );
  const apiVersion = botURL.slice(botURL.indexOf("api-version")).split("=")[1];
  const regionalChannelSettingsURL = `${environmentEndPoint}/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`;

  // Using refs instead of IDs to get the webchat and loading spinner elements
  const webChatRef = useRef<HTMLDivElement>(null);
  const loadingSpinnerRef = useRef<HTMLDivElement>(null);

  // A utility function that extracts the OAuthCard resource URI from the incoming activity or return undefined
  function getOAuthCardResourceUri(activity: any): string | undefined {
    const attachment = activity?.attachments?.[0];
    if (
      attachment?.contentType === "application/vnd.microsoft.card.oauth" &&
      attachment.content.tokenExchangeResource
    ) {
      return attachment.content.tokenExchangeResource.uri;
    }
  }

  const handleLayerDidMount = async () => {
    const MSALWrapperInstance = new MSALWrapper(
      props.clientID,
      props.authority
    );

    // Trying to get token if user is already signed-in
    let responseToken = await MSALWrapperInstance.handleLoggedInUser(
      [props.customScope],
      props.userEmail
    );

    if (!responseToken) {
      // Trying to get token if user is not signed-in
      responseToken = await MSALWrapperInstance.acquireAccessToken(
        [props.customScope],
        props.userEmail
      );
    }

    const token = responseToken?.accessToken || null;

    // Get the regional channel URL
    let regionalChannelURL;

    const regionalResponse = await fetch(regionalChannelSettingsURL);
    if (regionalResponse.ok) {
      const data = await regionalResponse.json();
      regionalChannelURL = data.channelUrlsById.directline;
    } else {
      console.error(`HTTP error! Status: ${regionalResponse.status}`);
    }

    // Create DirectLine object
    let directline: any;

    const token1 = localStorage.getItem("myToken");
    const isTokenExpired = (token: string | null): boolean => {
      if (!token) return true;

      const payload: { exp: number } = JSON.parse(atob(token.split(".")[1]));
      const currentTime: number = Math.floor(Date.now() / 1000);

      return payload.exp < currentTime;
    };
    if (token1 && !isTokenExpired(token1)) {
      directline = ReactWebChat.createDirectLine({
        token: token1,
        domain: regionalChannelURL + "v3/directline",
      });
    } else {
      const response = await fetch(botURL);
      const conversationInfo = await response.json();
      localStorage.setItem("myToken", conversationInfo.token);
      directline = ReactWebChat.createDirectLine({
        token: conversationInfo.token,
        domain: regionalChannelURL + "v3/directline",
      });
    }

    const store = ReactWebChat.createStore(
      {},
      ({ dispatch }: { dispatch: Dispatch }) =>
        (next: any) =>
        (action: any) => {
          // Checking whether we should greet the user
          //if (props.greet) {
            // if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
            //   console.log("Action:" + action.type);
            //   dispatch({
            //     meta: {
            //       method: "keyboard",
            //     },
            //     payload: {
            //       activity: {
            //         channelData: {
            //           postBack: true,
            //         },
            //         //Web Chat will show the 'Greeting' System Topic message which has a trigger-phrase 'hello'
            //         name: "startConversation",
            //         type: "event",
            //       },
            //     },
            //     type: "DIRECT_LINE/POST_ACTIVITY",
            //   });
            //   return next(action);
            // }
          //}
          if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
            dispatch({
              type: "WEB_CHAT/SEND_EVENT",
              payload: {
                name: "startConversation",
                type: "event",
                value: { text: "hello" },
              },
            });
            return next(action);
          }
          // Checking whether the bot is asking for authentication
          if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
            const activity = action.payload.activity;
            if (
              activity.from &&
              activity.from.role === "bot" &&
              getOAuthCardResourceUri(activity)
            ) {
              directline
                .postActivity({
                  type: "invoke",
                  name: "signin/tokenExchange",
                  value: {
                    id: activity.attachments[0].content.tokenExchangeResource
                      .id,
                    connectionName:
                      activity.attachments[0].content.connectionName,
                    token,
                  },
                  from: {
                    id: props.userEmail,
                    name: props.userFriendlyName,
                    role: "user",
                  },
                })
                .subscribe(
                  (id: any) => {
                    if (id === "retry") {

                      if (token1 && !isTokenExpired(token1)) {
                        console.log(
                          "since token already available, so do not display the oauthCard"
                        );
                        return;
                      }
                      // bot was not able to handle the invoke, so display the oauthCard (manual authentication)
                      console.log(
                        "bot was not able to handle the invoke, so display the oauthCard"
                      );
                      return next(action);
                    }
                  },
                  (error: any) => {
                    // an error occurred to display the oauthCard (manual authentication)
                    console.log("An error occurred so display the oauthCard");
                    return next(action);
                  }
                );
              // token exchange was successful, do not show OAuthCard
              return;
            }
          } else {
            return next(action);
          }

          return next(action);
        }
    );

    // hide the upload button - other style options can be added here
    const canvasStyleOptions = {
      hideUploadButton: true,
      botAvatarInitials: "B",
      userAvatarInitials: "U",
      botAvatarImage: botIcon,
      userAvatarImage: profileIcon,
    };

    // Render webchat
    if (token && directline) {
      if (webChatRef.current && loadingSpinnerRef.current) {
        loadingSpinnerRef.current.style.display = "none";
        ReactWebChat.renderWebChat(
          {
            directLine: directline,
            store: store,
            styleOptions: canvasStyleOptions,
            userID: props.userEmail,
          },
          webChatRef.current
        );
      } else {
        console.error("Webchat or loading spinner not found");
      }
    }
  };

  return (
    <div className="chat-window-wrapper">
      <div onClick={toggleHideDialog} className="chatButton">
        <Image
          src={fixedButtonIcon}
          width={50}
          height={50}
          alt="Chatbot Icon"
        />
      </div>

      <Dialog
        hidden={hideDialog}
        onDismiss={toggleHideDialog}
        onLayerDidMount={handleLayerDidMount}
        dialogContentProps={dialogContentProps}
        modalProps={modalProps}
        className="chat-window-wrapper"
      >
        <div className="chat-window">
          <div id="heading" className="chat-header">
            <div className="profile">
              <Image src={profileIcon} width={30} height={30} alt="" />
            </div>
            <h1>{props.botName}</h1>
            <span className="chatwindowCloseIcon" onClick={toggleHideDialog}>
              <Image src={closeIcon} width={16} height={16} alt="" />
            </span>
          </div>
          <div id="chatContainer" className="webchat">
            <div ref={webChatRef} role="main"></div>
            <div className="chat-loader" ref={loadingSpinnerRef}>
              {/* <Spinner label="Loading..." /> */}
                <Image
                  src={chatLoader}
                  width={50}
                  height={50}
                  alt="Chatbot Icon"
                />
            </div>
          </div>
        </div>
      </Dialog>
    </div>
  );
};

export default class Chatbot extends React.Component<IChatbotProps> {
  constructor(props: IChatbotProps) {
    super(props);
  }
  public render(): JSX.Element {
    return <PVAChatbotDialog {...this.props} />;
  }
}
