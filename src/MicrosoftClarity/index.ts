import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ClarityInterface, IClarityInterfaceProps, ClarityStatus } from "./components/ClarityInterface";
import * as React from "react";
import Clarity from '@microsoft/clarity';

export class MicrosoftClarity implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private notifyEventsPosted: () => void;
    private notifyTagsPosted: () => void;
    private onSelectHandler: () => void;
    private debounceTimeoutId: number | undefined;
    private debounceDelay: number = 200;

    private dispatchNotifyOutputChanged: boolean = false;
    private sessionPrioritized: boolean = false;
    private consentProvided: boolean = false;
    private clarityStatus: ClarityStatus = ClarityStatus.NotStarted;
    private currentProjectId: string | null = null;
    private currentSessionDetails: {
        userId: string,
        sessionId: string | undefined,
        activeScreenName: string | undefined,
        userName: string | undefined
    } = {
            userId: "",
            sessionId: "",
            activeScreenName: "",
            userName: ""
        };

    private lastPostedTags: any[] = [];

    constructor() { }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.notifyEventsPosted = context?.events?.OnEventPosting;
        this.notifyTagsPosted = context?.events?.OnTagsPosting;
        this.onSelectHandler = context?.events?.OnSelect;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.dispatchNotifyOutputChanged = false;

        /* check if both consent has been granted and a project id has been provided and/or changed
        to initialize Clarity if it hasn't already been initialized */
        const newConsentProvided = context.parameters.ConsentGranted.raw;
        const newProjectId = context.parameters.ClarityProjectId.raw;
        if (newConsentProvided !== this.consentProvided || newProjectId !== this.currentProjectId) {
            if (newConsentProvided) {
                this.initializeClarity(newProjectId as string);
            }
            this.consentProvided = newConsentProvided;
            this.dispatchNotifyOutputChanged = true;
        }

        /* check and prioritize session if it has been called for and hasn't already been prioritized */
        const newSessionPrioritized = context.parameters.PrioritizeSession.raw;
        if (newSessionPrioritized !== this.sessionPrioritized && newSessionPrioritized) {
            Clarity.upgrade("");
            this.sessionPrioritized = newSessionPrioritized;
            this.dispatchNotifyOutputChanged = true;
        }

        /* check if session details have changed and identify or re-identify the session if they have */
        const newSessionDetails = {
            userId: context.parameters.UserId.raw || "",
            sessionId: context.parameters.SessionId.raw || "",
            activeScreenName: context.parameters.ActiveScreenName.raw || "",
            userName: context.parameters.UserName.raw || ""
        };
        if (this.haveSessionDetailsChanged(newSessionDetails)) {
            this.identifySession(newSessionDetails);
            this.dispatchNotifyOutputChanged = true;
        }

        /* check if a new event message has been provided and post it if it is valid */
        const newEventMessage = context.parameters.CustomEvent.raw || "";
        if (this.isValidInput(newEventMessage, "Event message")) {
            this.postEvent(context, newEventMessage);
            this.dispatchNotifyOutputChanged = true;
        }

        /* check if new custom tags have been provided and post them if they are valid */
        const newCustomTags = context.parameters.CustomTags.raw || "";
        if (this.isValidJsonTags(newCustomTags)) {
            this.postTags(context, JSON.parse(newCustomTags));
            this.dispatchNotifyOutputChanged = true;
        }

        /* notify output changed if any of the above checks have resulted in a change */
        if (this.dispatchNotifyOutputChanged) {
            this.debouncedNotifyOutputChanged();
        }

        /* return the Clarity interface component to display */
        const props: IClarityInterfaceProps = {
            onSelect: this.onSelectHandler,
            clarityStatus: this.clarityStatus
        };
        return React.createElement(
            ClarityInterface, props
        );
    }

    public getOutputs(): IOutputs {
        return {
            ClarityStatus: ClarityStatus[this.clarityStatus],
        };
    }

    public destroy(): void { }

    private initializeClarity(projectId: string | undefined): void {
        /* check if project id is provided */
        if (!this.isValidInput(projectId, "Project ID")) {
            return;
        }

        /* check if Clarity has already been initialized with the same project id */
        if (projectId === this.currentProjectId && this.clarityStatus === ClarityStatus.Initialized) {
            return;
        }

        /* initialize and consent to Clarity with the provided project id */
        try {
            Clarity.init(projectId as string);
            Clarity.consent(true);
            this.currentProjectId = projectId as string;
            this.updateClarityStatus(ClarityStatus.Initialized);
        } catch (error) {
            this.updateClarityStatus(ClarityStatus.Error);
        }
    }

    private identifySession(sessionDetails: { userId: string, sessionId: string | undefined, activeScreenName: string | undefined, userName: string | undefined }): void {
        const { userId, sessionId, activeScreenName, userName } = sessionDetails;
        /* check if Clarity has been initialized */
        if (this.clarityStatus !== ClarityStatus.Initialized) {
            return;
        }
    
        /* check if user id has been provided/valid */
        if (!this.isValidInput(userId, "User ID")) {
            return;
        }
        
        /* identify the Clarity session with the provided details */
        try {
            Clarity.identify(userId, sessionId, activeScreenName, userName);
            this.currentSessionDetails = sessionDetails;
        } catch (error) {
            console.error("Clarity identify failed:", error);
        }
    }    

    private postEvent(context: ComponentFramework.Context<IInputs>, eventMessage: string): void {
        /* check if Clarity has been initialized */
        if (this.clarityStatus !== ClarityStatus.Initialized) {
            return;
        }

        /* check if event message is provided/valid */
        if (!this.isValidInput(eventMessage, "Event message")) {
            return;
        }

        /* post the event message to Clarity */
        try {
            Clarity.event(eventMessage);
            this.notifyEventsPosted();
        } catch (error) {
            console.error("Clarity event failed:", error);
        }
    }

    private postTags(context: ComponentFramework.Context<IInputs>, tags: any[]): void {
        /* check if Clarity has been initialized */
        if (this.clarityStatus !== ClarityStatus.Initialized) {
            return;
        }

        /* check if the session tags are have changed from previously posted tags */
        if (JSON.stringify(tags) === JSON.stringify(this.lastPostedTags)) {
            return;
        }

        /* post the session tags to Clarity */
        try {
            tags.forEach(tag => {
                if (tag.Key && tag.Value) {
                    Clarity.setTag(tag.Key, tag.Value);
                }
            });
            this.lastPostedTags = tags;
            this.notifyTagsPosted();
        } catch (error) {
            console.error("Clarity tag posting failed:", error);
        }
    }

    private haveSessionDetailsChanged(newSessionDetails: { userId: string, sessionId: string | undefined, activeScreenName: string | undefined, userName: string | undefined }): boolean {
        return newSessionDetails.userId !== this.currentSessionDetails.userId ||
            newSessionDetails.sessionId !== this.currentSessionDetails.sessionId ||
            newSessionDetails.activeScreenName !== this.currentSessionDetails.activeScreenName ||
            newSessionDetails.userName !== this.currentSessionDetails.userName;
    }

    private isValidInput(value: string | undefined | null, type: string): boolean {
        if (!value || value.trim() === "" || value === "val") {
            return false;
        }
        return true;
    }

    private isValidJsonTags(value: string): boolean {
        /* check if value is provided */
        if (!value || value.trim() === "") {
            return false;
        }

        /* check if value is a valid JSON array of tags */
        try {
            const parsed = JSON.parse(value);
            if (!Array.isArray(parsed) || parsed.length === 0) {
                return false;
            }
            return true;
        } catch (error) {
            console.error("Invalid JSON format for custom tags:", error);
            return false;
        }
    }

    private updateClarityStatus(status: ClarityStatus): void {
        this.clarityStatus = status;
    }

    private debouncedNotifyOutputChanged(): void {
        if (this.debounceTimeoutId) {
            clearTimeout(this.debounceTimeoutId);
        }
        this.debounceTimeoutId = window.setTimeout(() => {
            this.notifyOutputChanged();
        }, this.debounceDelay);
    }
}