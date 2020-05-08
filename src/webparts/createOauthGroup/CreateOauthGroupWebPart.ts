import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import * as strings from 'CreateOauthGroupWebPartStrings';

import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import 'jquery';
import { MSGraphClient } from '@microsoft/sp-http';

import '../../ExternalRef/css/alertify.min.css';
import '../../ExternalRef/css/style.css';
var alertify: any = require('../../ExternalRef/js/alertify.min.js');
SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");
declare var $;
var globalDisplayName = '';
var createPlannerNow = false;

export interface ICreateOauthGroupWebPartProps {
  description: string;
}

export default class CreateOauthGroupWebPart extends BaseClientSideWebPart<ICreateOauthGroupWebPartProps> {


  public render(): void {

    this.domElement.innerHTML = `

      <div><h1>Create AD Group</h1><br/>
      <div class="form-group">
      
      <div class="row row-10-gap">
      <div class="col-sm-2">
      <label>Group Name: </label>
      <select id="countryCode" class="form-control">
        <option value="USA">USA</option>
        <option value="Brazil">Brazil</option>
      </select>
      </div>
      
      <div class="col-sm-2">
      <label>Company Code: </label>
      <select id="companyCode" class="form-control">
        <option value="AZT">AZT</option>
      </select> 
      </div>
      <div class="col-sm-2">
      <label>Group Code: </label>
      <select id="groupCode" class="form-control" >
        <option value="AABBB">AABBB</option>
      </select> 
      </div>
      <div class="col-sm-2">
      <label>Project Number: </label>
      <select id="projectNumber" class="form-control">
        <option value="QWERT">QWERT</option>
      </select> 
      </div>
      <div class="col-sm-2">
      <label>Task Number:</label>
      <input type="text" id="taskNumber" name="taskNumber" class="form-control"> 
      </div>
      <div class="col-sm-2">
      <label style="white-space: nowrap;">Short Description: </label>
      <input type="text" id="shortDescription" name="shortDescription" class="form-control" style="width: 120px;"> 
      </div>
      </div>
</div>

      <div class="form-group">
      <label>Members: </label>
      <input type="text" id="members" name="members" class="form-control" />
      </div>

      


      <div class="form-group">
      <label>Description: </label>
      <textarea id="description" rows="4" cols="50" class="form-control"></textarea>
</div>
<div class="row">
<div class="col-sm-4">
<div class="form-group">
      <label>Visibility: </label>
      <select id="visibility" class="form-control">
        <option value="Public">Public</option>
        <option value="Private">Private</option>
      </select> 
      </div>
      </div>
      <div class="col-sm-4">
<div class="form-group" style="margin-top: 30px;margin-left: -15px;">
      
      <input type="checkbox" id="sendmail" name="sendmail" class="radio-stylish">
      <span class="checkbox-element"></span>
      <label for="sendmail" class="stylish-label">Send mail to members: </label> 
</div>
</div>
</div>
<div class="form-group">
      <input type="button" id="btnSubmit" value="Submit" class="btn btn-primary">
      
      </div>
      </div>`;

    var that = this;
    $('#btnSubmit').on('click', function (e) {

      alertify.set('notifier', 'position', 'top-right');

      if (!$('#countryCode').val()) {
        alertify.error('Country code is required');
        return;
      }

      if (!$('#companyCode').val()) {
        alertify.error('Company code is required');
        return;
      }

      if (!$('#groupCode').val()) {
        alertify.error('Group code is required');
        return;
      }

      if (!$('#projectNumber').val()) {
        alertify.error('Project number is required');
        return;
      }

      if (!$('#taskNumber').val()) {
        alertify.error('Task number is required');
        return;
      }

      if (!$('#shortDescription').val()) {
        alertify.error('Short description is required');
        return;
      }

      that.createADGroup();
    });
  }

  addmembertogroup(userid, email, groupId) {
    var that = this;
    var user = {
      "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + userid
    };
    this.context.msGraphClientFactory.getClient().then(
      (client: MSGraphClient): void => {
        client
          .api('/groups/' + groupId + '/members/$ref')
          .post(user)
          .then((content: any) => {
            // alertify.success('User ' + email + ' added');

            if (createPlannerNow) {
              createPlannerNow = false;
              setTimeout(function () {
                that.createPlanner(groupId, globalDisplayName);
                var postData = {
                  "groupId": groupId
                };
                $.ajax({
                  url: 'https://prod-09.centralindia.logic.azure.com:443/workflows/54949b246a8645069ba77a35619daa08/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=3UodVAI2oglFIB9tZq3-HubBvnrBmd_yxfLkPVWKFxU',
                  type: "POST",
                  data: JSON.stringify(postData),
                  contentType: "application/json; charset=utf-8",
                  success: function (data) {
                    var delay = alertify.get("notifier", "delay");
                    alertify.set("notifier", "delay", 10);
                    alertify.set("notifier", "position", "top-right");
                    alertify
                      .alert(
                        "Data Processed sucessfully",
                        function() {
                          location.reload();
                        }
                      )
                      .setHeader("<em> Saved </em> ");
                //  alertify.success('Teams created successfully');
                  }, error: function (err) {
                    // alertify.success('Error while creating team');
                  }
                });

              }, 5000)
            }

          })
          .catch(err => {
            // alertify.error('Error while creating member');
          });
      }
    );
  }

  getuser(email, groupId) {
    var that = this;
    var cleartext = email.replace(/\s+/g, '');
    this.context.msGraphClientFactory.getClient().then(
      (client: MSGraphClient): void => {
        client
          .api('/users/' + cleartext)
          .get()
          .then((content: any) => {
            that.addmembertogroup(content.id, email, groupId);
          })
          .catch(err => {

          });
      }
    );
  }

  createADMembers(groupId) {
    var members = $('#members').val().split(';');
    if (members.length == 0) {
      createPlannerNow = true;
    }
    this.getuser(this.context.pageContext.user.email, groupId);
    for (let index = 0; index < members.length; index++) {
      const member = members[index];
      if (index == members.length - 1) {
        createPlannerNow = true;
      }
      if (member) {
        this.getuser(member, groupId);
      }
    }
  }

  createPlanner(groupId, title) {
    var plannerPlan = {
      owner: groupId,
      title: title
    };
    this.context.msGraphClientFactory.getClient().then(
      (client: MSGraphClient): void => {
        client
          .api('/planner/plans')
          .post(plannerPlan)
          .then((content: any) => {
            // alertify.success('Planner created successfully');
          })
          .catch(err => {
            // alertify.error('Error while creating planner');
          });
      }
    );
  }

  // createTeams(groupId) {
  //   var team = {
  //     memberSettings: {
  //       allowCreateUpdateChannels: true
  //     },
  //     messagingSettings: {
  //       allowUserEditMessages: true,
  //       allowUserDeleteMessages: true
  //     },
  //     funSettings: {
  //       allowGiphy: true,
  //       giphyContentRating: "strict"
  //     }
  //   };
  //   this.context.msGraphClientFactory.getClient().then(
  //     (client: MSGraphClient): void => {
  //       client
  //         .api('/groups/' + groupId + '/team')
  //         .put(team)
  //         .then((content: any) => {
  //           alertify.success('Teams created successfully');
  //         })
  //         .catch(err => {
  //           alertify.error('Error while creating teams');
  //         });
  //     }
  //   );
  // }

  createADGroup() {
    var mailNickname = $('#countryCode').val() + '-' + $('#companyCode').val() + '-' + $('#groupCode').val() + '-' + $('#projectNumber').val() + '-' + $('#taskNumber').val();
    var displayName = mailNickname + '-' + $('#shortDescription').val();
    var clearMailNickname = mailNickname.replace(/[^a-zA-Z0-9]/g, "");


    globalDisplayName = displayName;

    var details = {
      "displayName": displayName,
      "groupTypes": [
        "Unified"
      ],
      "mailEnabled": true,
      "mailNickname": clearMailNickname,
      "securityEnabled": false,
      "visibility": $('#visibility').val()
    };

    if ($('#description').val()) {
      details["description"] = $('#description').val();
    }

    var that = this;
    this.context.msGraphClientFactory.getClient().then(
      (client: MSGraphClient): void => {
        client
          .api('/groups')
          .post(details)
          .then((content: any) => {
            that.createADMembers(content.id);
            //that.createTeams(content.id);
            // alertify.success('Group created successfully');
          })
          .catch(err => {
            // alertify.error('Error while creating a group');
          });
      }
    );

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
  }
}
