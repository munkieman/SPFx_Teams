var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './SpFxTeamsTogetherWebPart.module.scss';
import * as strings from 'SpFxTeamsTogetherWebPartStrings';
var SpFxTeamsTogetherWebPart = /** @class */ (function (_super) {
    __extends(SpFxTeamsTogetherWebPart, _super);
    function SpFxTeamsTogetherWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    SpFxTeamsTogetherWebPart.prototype.render = function () {
        var title = (this.teamsContext)
            ? 'Teams'
            : 'SharePoint';
        var currentLocation = (this.teamsContext)
            ? "Team: " + this.teamsContext.teamName
            : "site collection " + this.context.pageContext.web.title;
        this.domElement.innerHTML = "\n      <div class=\"" + styles.spFxTeamsTogether + "\">\n        <div class=\"" + styles.container + "\">\n          <div class=\"" + styles.row + "\">\n            <div class=\"" + styles.column + "\">\n              <span class=\"" + styles.title + "\">Welcome to " + title + "!</span>\n              <p class=\"" + styles.subTitle + "\">Currently in the context of the following " + currentLocation + "</p>\n              <p class=\"" + styles.description + "\">" + escape(this.properties.description) + "</p>\n              <a href=\"https://aka.ms/spfx\" class=\"" + styles.button + "\">\n                <span class=\"" + styles.label + "\">Learn more</span>\n              </a>\n            </div>\n          </div>\n        </div>\n      </div>";
    };
    SpFxTeamsTogetherWebPart.prototype.onInit = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (_this.context.sdks.microsoftTeams) {
                _this.teamsContext = _this.context.sdks.microsoftTeams.context;
            }
            resolve();
        });
    };
    Object.defineProperty(SpFxTeamsTogetherWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    SpFxTeamsTogetherWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return SpFxTeamsTogetherWebPart;
}(BaseClientSideWebPart));
export default SpFxTeamsTogetherWebPart;
//# sourceMappingURL=SpFxTeamsTogetherWebPart.js.map