# Microsoft Clarity PCF Control

## Overview

This project is a Power Platform PCF (PowerApps Component Framework) control that allows users to record and upload sessions to Microsoft Clarity utilizing the Microsoft Clarity API. It provides a seamless integration with Microsoft Clarity, enabling enhanced session insights and analytics within your Power Platform applications.

> Learn more about Microsoft Clarity at [clarity.microsoft.com](https://clarity.microsoft.com/)

## Features

- **Session Recording**: Capture and upload user sessions to your Microsoft Clarity project.
- **Custom Events**: Post custom events to Clarity with ease.
- **Session Tagging**: Add custom tags to sessions for better categorization and analysis.
- **Consent Management**: Ensures sessions are not captured unless and until you have granted or obtained consent from your users.
- **Prioritization**: Optionally prioritize sessions for recording and retention.
- **Seamless Integration**: Built with Microsoft technologies for easy integration.

## Installation

The latest solution files are automatically generated and made available via GitHub Releases. To install the component:

1. Go to the [Releases](https://github.com/ohthreesixtyfive/microsoft-clarity/releases) page of this repository.
2. Download the latest **managed** release solution file.
3. Import the solution file into your Power Platform environment.
4. Create or open a canvas app or custom page.
5. From the side bar **Insert** menu, select the **Get more components** icon.
6. From the **Import components** side bar, navigate to the **Code** tab.
7. Select `Microsoft Clarity` from the list of components and then select **Import**.
8. From the side bar **Insert** menu, expand **Code components** and select the `Microsoft Clarity` component to place it on your canvas.

> [!NOTE]
> Code components for Power Apps must be enabled in your environment before use.
> 
> Learn more: [Microsoft Learn | Power Apps: Enable the Power Apps component framework feature](https://learn.microsoft.com/en-us/power-apps/developer/component-framework/component-framework-for-canvas-apps#enable-the-power-apps-component-framework-feature)

## Properties
- **ClarityProjectId** [Required]: The ID of the Clarity project used to log events and sessions.
- **ConsentGranted** [Required]: Boolean indicating if consent has been granted for recording and tracking the current user session.
- **SessionId**: The PowerApps session ID. Optional, defaults to `Host.SessionID`.
- **ActiveScreenName**: The name of the active screen within the app. Optional, defaults to `App.ActiveScreen.Name`.
- **UserId**: The user's ID, typically an email address. Optional, defaults to `User().Email`.
- **UserName**: The full name of the user. Optional, defaults to `User().FullName`.
- **PrioritizeSession**: Boolean to prioritize sessions for recording, defaults to `false`.

> [!NOTE]
> At minimum, both a `ClarityProjectId` and setting `ConsentGranted` to *true* is required to initialize recording of a session.

## Advanced Features

### Customizing Session Insights with Events, Tags, and Custom Identifiers

Enhance your analytics by leveraging advanced features such as custom events, session tags, and custom identifiers. These capabilities allow for a deeper understanding and more precise tracking within your Microsoft Clarity implementation.

> [!NOTE]
> Custom events and identifiers may not immediately appear in session recordings and could take several hours to display or be available for filtering.

**Custom Identifiers**:
- Use custom identifiers like `UserId` and `SessionId` to provide consistent identification of users and sessions across different devices and browsers. This helps in mapping a user's journey across devices, which Clarity's default identifiers might not capture.
- Setting a `UserId` (such as an email) allows you to track a user's journey across different sessions and devices, providing a comprehensive view of the user experience.

> [!IMPORTANT]
> If a value for `SessionId` is not provided, the session **will not** be identified in Microsoft Clarity.

> [!NOTE]
> Clarity does not store custom identifiers as plain text. Instead, they are hashed on the client before being sent to the servers.
> 
> When you filter on a specific custom user ID, Clarity hashes the input and matches it against stored data to retrieve the right sessions.
> 
> Learn more here: [Microsoft Learn | Microsoft Clarity: Storing and managing Custom IDs](https://learn.microsoft.com/en-us/clarity/setup-and-installation/identify-api#storing-and-managing-custom-ids)

**Post a Custom Event**:
- Custom events are specific actions or occurrences that you want to track within user sessions. These could be button clicks, form submissions, or any other user interactions that are important for your analysis.
- Use the context variable `_ClarityEventMessage` to send an event message. Assign your custom event message like this:

  ```javascript
  UpdateContext({_ClarityEventMessage: "{Your Event Message}"});
  ```
> [!NOTE]
> When an event message is posted, its corresponding context variable will automatically be cleared.

> [!NOTE]
> You can post multiple events multiple times per session. Each event is logged individually and can be filtered or viewed across all verticals.
> 
> Learn more here: [Microsoft Learn | Microsoft Clarity: Smart Events](https://learn.microsoft.com/en-us/clarity/setup-and-installation/smart-events)

**Post Session Tags**:
- You can assign labels to sessions to categorize them based on specific criteria, such as user type, current environment, session purpose, or any other custom attributes that are relevant to your analysis.
- Use the collection `_ClaritySessionTags` to add tags to a session. Add your custom key-value pairs like this:

    ```javascript
    Collect(
      _ClaritySessionTags,
      {
        Key: "{CustomKey}",
        Value: "{CustomValue}"
      }
    );
    ```

**Session Prioritization**:
  - You may set **PrioritizeSession** to `true` to ensure that specific sessions are recorded and retained, especially when your total session volume exceeds Clarity's maximum daily per-project limit.
  - This is especially useful for focusing on sessions with important events or interactions.
  
> [!NOTE]
> Learn more about this feature and any applicable limits here: [Microsoft Learn | Microsoft Clarity: Session Prioritization](https://learn.microsoft.com/en-us/clarity/setup-and-installation/clarity-api#prioritize-specific-sessions-for-recording)


## License

This project is licensed under the MIT License.