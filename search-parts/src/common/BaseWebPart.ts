import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
    PropertyPaneLabel,
    PropertyPaneLink
} from '@microsoft/sp-property-pane';
import IExtensibilityService from '../services/extensibilityService/IExtensibilityService';
import { ExtensibilityService } from '../services/extensibilityService/ExtensibilityService';
import { IBaseWebPartProps } from "../models/common/IBaseWebPartProps";
import * as commonStrings from 'CommonStrings';
import { ThemeProvider, IReadonlyTheme, ThemeChangedEventArgs } from '@microsoft/sp-component-base';
import { isEqual } from '@microsoft/sp-lodash-subset';
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';

/**
 * Generic abstract class for all Web Parts in the solution
 */
export abstract class BaseWebPart<T extends IBaseWebPartProps> extends BaseClientSideWebPart<IBaseWebPartProps> {

    /**
     * Theme variables
     */
    protected _themeProvider: ThemeProvider;
    protected _themeVariant: IReadonlyTheme;

    /**
     * The Web Part properties
     */
    protected properties: T;

    /**
     * The data source service instance
     */
    protected extensibilityService: IExtensibilityService = undefined;

    constructor() {
        super();
    } 
    
    /**
     * Initializes services shared by all Web Parts. Need to be called in the consumer onInit() method.
     */
    protected async initializeSharedFeaturesAndServices(): Promise<void> {

        // Get service instances
        this.extensibilityService = this.context.serviceScope.consume<IExtensibilityService>(ExtensibilityService.ServiceKey);

        // Initializes them variant
        this.initThemeVariant();
        
        return;
    }

    /**
     * Returns common information groups for the property pane
     */
    protected getPropertyPaneWebPartInfoGroups() {

        return [
            {
                groupName: commonStrings.General.About,
                groupFields: [
                  PropertyPaneWebPartInformation({
                    description: `<span>${commonStrings.General.Authors}: Franck Cornu <a href="https://www.linkedin.com/in/franckcornu/" target="_blank"><img width="16px" src="https://static-exp1.licdn.com/sc/h/79m1qiu8wnezlizb5zn646ata"/></a> <a href="https://twitter.com/FranckCornu"><img width="16px" src="https://upload.wikimedia.org/wikipedia/fr/thumb/c/c8/Twitter_Bird.svg/1200px-Twitter_Bird.svg.png"/></a></span>`,
                    key: 'authors'
                  }),
                  PropertyPaneLabel('', {
                    text: `${commonStrings.General.Version}: ${this && this.manifest.version ? this.manifest.version : ''}`
                  }),   
                  PropertyPaneLabel('', {
                    text: `${commonStrings.General.InstanceId}: ${this.instanceId}`
                  }),            
                ]
              },
              {
                groupName: commonStrings.General.Resources.GroupName,
                groupFields: [
                  PropertyPaneLink('',{
                    target: '_blank',
                    href: this.properties.documentationLink,
                    text: commonStrings.General.Resources.Documentation
                  })
                ]
              }
        ];
    }


    /**
     * Initializes theme variant properties
     */
    private initThemeVariant(): void {

        // Consume the new ThemeProvider service
        this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

        // If it exists, get the theme variant
        this._themeVariant = this._themeProvider.tryGetTheme();

        // Register a handler to be notified if the theme variant changes
        this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent.bind(this));
    }

    /**
     * Update the current theme variant reference and re-render.
     * @param args The new theme
     */
    private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {

        if (!isEqual(this._themeVariant, args.theme)) {
            this._themeVariant = args.theme;
            this.render();
        }
    }
}