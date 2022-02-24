import {IInputs, IOutputs} from "./generated/ManifestTypes";
import React from 'react';
import ReactDOM from 'react-dom';
//import './index.css';
import App from './App';
//import * as serviceWorker from './serviceWorker';

import { Providers } from '@microsoft/mgt-element';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';

export class PCFwithMGT implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// The PCF context object
	private context: ComponentFramework.Context<IInputs>;

	// The wrapper div element for the component
	private container: HTMLDivElement;

	// The callback function to call whenever your code has made a change to a bound or output property
	private notifyOutputChanged: () => void;

	// Flag to track if the component is in edit mode or not
	private isEditMode: boolean;

	// Tracks the event handler so we can destroy it when done
	private buttonClickHandler: EventListener;

	// Tracking variable for the name property
	private name: string | null;

	private redirectUrl: string | null;
	private clientId: string | null;
	private graphScopes: string[] | null;
	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
	{
		// Track all the things
		this.context = context;
		this.notifyOutputChanged = notifyOutputChanged;
		this.container = container;
		this.isEditMode = false;
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.name = context.parameters.Name.raw;
		this.redirectUrl = context.parameters.RedirectUrl.raw;
		this.clientId = context.parameters.ClientId.raw;
		this.graphScopes = this.getScopes(context.parameters.GraphScopes.raw??"'calendars.read', 'user.read', 'openid', 'profile', 'people.read', 'user.readbasic.all', 'files.read', 'files.read.all'");
		Providers.globalProvider = new Msal2Provider({
			clientId: this.clientId??'5d442b28-b1ff-49bc-b51a-a2c5b7e122df', //'ba686da8-8cb8-4e41-9765-056a10dee34c',
			scopes: this.graphScopes ?? ['calendars.read', 'user.read', 'openid', 'profile', 'people.read', 'user.readbasic.all', 'files.read', 'files.read.all'],
			redirectUri: this.redirectUrl??'https://apps.powerapps.com'
		  });

		ReactDOM.render(
		// Create the React component
		React.createElement(
			App
		),
		this.container
		);
	}

	private getScopes = (scopesCSV: string) => {
		let scopes: string[] = [];
		if (scopesCSV) {
			scopes = scopesCSV
						.split(",")
						.map(s => s.trim());
		}
		return scopes;
	}
	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/**
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}
}
