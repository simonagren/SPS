import { DisplayMode } from "@microsoft/sp-core-library";

export interface ISpFxProvisioningProps {
  displayMode: DisplayMode;
  listId: string;
  onConfigure: () => void;
  siteUrl?: string;
  /**
   * Web part title to show in the body
   */
  title: string;
  /**
   * Event handler after updating the web part title
   */
  updateProperty: (value: string) => void;
}
