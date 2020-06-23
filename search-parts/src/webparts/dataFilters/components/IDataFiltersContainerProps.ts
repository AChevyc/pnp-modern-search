import { IDataFilterConfiguration, IDataFilter, IDataFilterResult } from "@pnp/modern-search-extensibility";
import IDataFiltersWebPartProps from "../IDataFiltersWebPartProps";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ITemplateService } from "../../../services/templateService/ITemplateService";
import { IWebPartTitleProps } from "@pnp/spfx-controls-react/lib/WebPartTitle";

export interface IDataFiltersContainerProps {

  /**
   * The template content before processing
   */
  templateContent: string;

  /**
   * The filters to display
   */
  availableFilters: IDataFilterResult[];

  /**
   * The filters configuration
   */
  filtersConfiguration: IDataFilterConfiguration[];

  /**
   * The Web Part DOM element
   */
  domElement: HTMLElement;

  /**
   * The Web Part instance ID
   */
  instanceId: string;

  /**
   * The selected layout key
   */
  selectedLayoutKey: string;

  /**
   * The Web Part properties so they can be used in Handlebars template
   */
  properties: IDataFiltersWebPartProps;

  /**
   * The current theme information
   */
  themeVariant: IReadonlyTheme;

  /**
   * Handler method when a filter value is updated in children components
   */
  onUpdateFilters: (filters: IDataFilter[]) => void;

  /**
   * A template service instance
   */
  templateService: ITemplateService;

  /**
   * The Web Part Title props
   */
  webPartTitleProps: IWebPartTitleProps;
}
