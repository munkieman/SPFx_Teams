import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface ISpFxTeamsTogetherWebPartProps {
    description: string;
}
export default class SpFxTeamsTogetherWebPart extends BaseClientSideWebPart<ISpFxTeamsTogetherWebPartProps> {
    render(): void;
    private teamsContext;
    protected onInit(): Promise<void>;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=SpFxTeamsTogetherWebPart.d.ts.map