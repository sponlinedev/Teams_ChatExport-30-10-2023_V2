import * as React from 'react';
import styles from './TeamsChatExportWp.module.scss';
import { ITeamsChatExportWpProps } from './ITeamsChatExportWpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient, HttpClientResponse } from '@microsoft/sp-http';
import { JsonToTable } from "react-json-to-table";
import { TextField, MaskedTextField } from '@fluentui/react/lib/TextField';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import {DetailsListDocumentsExample} from './ResultTable';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import ReactHtmlParser from "react-html-parser";
import html2pdf from 'html2pdf.js';

const options = {
  orientation: 'landscape',
  unit: 'in',
  format: [4,2]
};

export default class TeamsChatExportWp extends React.Component<any, any> {
  constructor(props: any) {
    super(props);
    

    // Initialize the state of the component
    this.state = {
      teamExportPackage: [],
      allTeams: [],
      allChannels: [],
      selectedTeamId: "",
      selectedTeamName: "",
      selectedChannelId: "",
      selectedChannelName: "",
      isChatClicked: "none",
      sampleText: "This is sample text",
      _allItems: [],
      isCsvVisible: "none",
      isPdfVisible: "none",
      fromDate: null,
      toDate: null,
      topload: [{ text: 10 , key: 0 },{ text: 20 , key: 1 },{ text: 50 , key: 2 },{ text: 100 , key: 3 }],
      _topLoadSelect: 10,
      defaultSelectedKey: 0
    };
  }

componentDidMount() {
  this.getMyTeams();
}

private async getMyTeams()
{
    var _graphClient = await this.props.context.msGraphClientFactory.getClient();
  
    const teamsResponse = await _graphClient.api('me/joinedTeams').version('v1.0').get();  
    var myTeams = teamsResponse.value as [];  
    var _tempTeams = [];

    await myTeams.map(async (_myTeam: any) => {
      _tempTeams.push({ text: _myTeam.displayName , key: _myTeam.id });
    })

    this.setState({
      allTeams: _tempTeams
    })
}

private _onTeamsDropDownChanged(event) { 
    this.setState({
      selectedTeamId: event.key,
      selectedTeamName: event.text,
      isChatClicked: "none"
    })
  this.getChannels(event.key)
}

private async getChannels(_selectedTeamKey)
{
    var _graphClient = await this.props.context.msGraphClientFactory.getClient();
  
    const channelsResponse = await _graphClient.api('teams/' + _selectedTeamKey + '/channels').version('v1.0').get();  
    var myChannels = channelsResponse.value as []; 
    var _tempChannels = [];
    
    myChannels.map(async (_myChannel: any) => {          
      _tempChannels.push({ text: _myChannel.displayName , key: _myChannel.id });
    })

    this.setState({
      allChannels: _tempChannels
    })

    this._onChannelDropDownChanged(_tempChannels[0])
}

private _onChannelDropDownChanged(event) { 
  this.setState({
    selectedChannelId: event.key,
    selectedChannelName: event.text,
    isChatClicked: "none"
  })  
}


private _onTopLoadDropDownChanged(event) { 
  this.setState({
    _topLoadSelect: event.text,
    defaultSelectedKey: event.key,
  })  
}

private csvButtonClicked = async () => {

  this.setState({
    isCsvVisible: "block",
    isPdfVisible: "none"
  })

  this.exportMessage();
}

private pdfButtonClicked = async () => {

  this.setState({
    isCsvVisible: "none",
    isPdfVisible: "block"
  })

  this.exportMessage();
}

private async exportMessage()
{
  console.log(this.state.selectedTeamName + " - " + this.state.selectedChannelName);
  var _graphClient = await this.props.context.msGraphClientFactory.getClient(); 
  var _resultMessages = [], _resultReplies = [];
  let messages, replies; 
  var varDateFrom = new Date(this.state.fromDate);
  var varDateTo = new Date(this.state.toDate);  

  let promise = new Promise(async (resolve, reject) => {
    const messagesResponse = await _graphClient.api('teams/' + this.state.selectedTeamId + '/channels/' + this.state.selectedChannelId + '/messages').version('v1.0').get();  
    messages = messagesResponse.value as [];  

    messages.map(async (_message: any) => {  

      if(_message.from)
      {
        var _messageDate = new Date(_message.lastModifiedDateTime);

        if(this.state.fromDate && this.state.toDate)
        {
          if(varDateFrom <= _messageDate && varDateTo >= _messageDate)
          {
            _resultMessages.push({"messageId": _message.id, "messageDateTime": _message.lastModifiedDateTime ,"messageBy": _message.from.user.displayName, "messageContent": _message.body.content})
          }
        }
        else
        {
          _resultMessages.push({"messageId": _message.id, "messageDateTime": _message.lastModifiedDateTime ,"messageBy": _message.from.user.displayName, "messageContent": _message.body.content})          
        }
      }

      try {  
        const repliesResponse = await _graphClient.api('teams/' + this.state.selectedTeamId + '/channels/' + this.state.selectedChannelId + '/messages/' + _message.id + '/replies').version('v1.0').get();  
        replies = repliesResponse.value as [];  

        if(replies.length > 0) 
        {
          replies.map(async (_reply: any) => {  
            if(_reply.from)
            {
              var _replyDate = new Date(_reply.lastModifiedDateTime);

              if(this.state.fromDate && this.state.toDate)
              {
                if(varDateFrom <= _replyDate && varDateTo >= _replyDate)
                {
                  _resultReplies.push({"messageId": _message.id,"replyId": _reply.id, "replyToId": _reply.replyToId, "messageDateTime": _reply.lastModifiedDateTime ,"messageBy": _reply.from.user.displayName, "messageContent": _reply.body.content})
                }
              }
              else
              {
                _resultReplies.push({"messageId": _message.id,"replyId": _reply.id, "replyToId": _reply.replyToId, "messageDateTime": _reply.lastModifiedDateTime ,"messageBy": _reply.from.user.displayName, "messageContent": _reply.body.content})                
              }

              setTimeout(() => resolve("done!"), 1000);
            }
            ///
          })                    
        }
        else
        {
          setTimeout(() => resolve("done!"), 1000);
        }
      } catch (error) {  
        console.log('unable to get channels', error);  
      }
    })    
  });

  let result = await promise; // wait until the promise resolves (*)

  console.log("Result...");
  console.log(result);
  
  var _finalResult = [];
  var _varMessageSrNo = 0;
  var _masterSrNo = 0;
  var _childSrNo = 0;

  _resultMessages.map(async (message: any) => {
    _varMessageSrNo++;
    _childSrNo = 0;
    var _date = new Date(message.messageDateTime);
    var _month = parseInt(_date.getMonth().toString())+1;
    var _messageFormattedDate = _date.getDate() + "-"+ _month + "-" + _date.getFullYear() + " " + _date.getHours() + ":" + _date.getMinutes() + ":" + _date.getSeconds();
    var _messaageId = _varMessageSrNo;
    _finalResult.push({
      "ID": _messaageId,
      "ReplyID": 0,
      "Message": ReactHtmlParser(message.messageContent),
      "Sender": message.messageBy,
      "DateTime": _messageFormattedDate,  
      "MessageType": "Main",
      "OriginalDate": _date,
      "masterSrNo": _masterSrNo,
      "childSrNo": _childSrNo,
      "MessageCSV": message.messageContent,
    });

    /////////////
    _resultReplies.map(async (reply: any) => {
        if(message.messageId == reply.replyToId)
        {
          var _date = new Date(reply.messageDateTime);
          var _month = parseInt(_date.getMonth().toString())+1;
          var _replyFormattedDate = _date.getDate() + "-"+ _month + "-" + _date.getFullYear() + " " + _date.getHours() + ":" + _date.getMinutes() + ":" + _date.getSeconds();
          _varMessageSrNo++;
          _childSrNo++;

          _finalResult.push({
            "ID": _varMessageSrNo,
            "ReplyID": _messaageId,
            "Message": ReactHtmlParser(reply.messageContent),
            "Sender": reply.messageBy,
            "DateTime": _replyFormattedDate,
            "MessageType": "Reply", 
            "OriginalDate": _date,
            "masterSrNo": _masterSrNo,
            "childSrNo": _childSrNo,
            "MessageCSV": message.messageContent
          })
        }
    });
    /////////////
    _masterSrNo++;
})

  console.log("Final Result...");
  console.log(_finalResult);

  this.setState((prevState) => ({
    isChatClicked: "block",
    _allItems: _finalResult.slice(0, prevState._topLoadSelect),
    teamExportPackage: _finalResult.slice(0, prevState._topLoadSelect),
    fromDate: null,
    toDate: null,
    defaultSelectedKey: 0
  }));  
}

private async getTeamsDetail()
{
    var _graphClient = await this.props.context.msGraphClientFactory.getClient(); //TODO  
    let myTeams, myChannels, messages, replies;
    var _resultMessages = [], _resultReplies = [];

    /////
    let promise = new Promise(async (resolve, reject) => {
      try {

        const teamsResponse = await _graphClient.api('me/joinedTeams').version('v1.0').get();  
        myTeams = teamsResponse.value as [];  

        await myTeams.map(async (_myTeam: any) => {
          try
          {
            const channelsResponse = await _graphClient.api('teams/' + _myTeam.id + '/channels').version('v1.0').get();  
            myChannels = channelsResponse.value as []; 
            
            myChannels.map(async (_myChannel: any) => {          
              try {  

                const messagesResponse = await _graphClient.api('teams/' + _myTeam.id + '/channels/' + _myChannel.id + '/messages').version('v1.0').get();  
                messages = messagesResponse.value as [];  

                messages.map(async (_message: any) => {  

                  if(_message.from)
                  {
                    _resultMessages.push({"teamId": _myTeam.id, "channelId": _myChannel.id,  "messageId": _message.id, "teamName": _myTeam.displayName, "channelName": _myChannel.displayName ,"messageDateTime": _message.createdDateTime ,"messageBy": _message.from.user.displayName, "messageContent": _message.body.content})
                  }

                  try {  
                    const repliesResponse = await _graphClient.api('teams/' + _myTeam.id + '/channels/' + _myChannel.id + '/messages/' + _message.id + '/replies').version('v1.0').get();  
                    replies = repliesResponse.value as [];  

                    if(replies.length > 0) 
                    {
                      replies.map(async (_reply: any) => {  
                        if(_reply.from)
                        {
                          _resultReplies.push({"teamId": _myTeam.id, "channelId": _myChannel.id, "messageId": _message.id,"replyId": _reply.id, "replyToId": _reply.replyToId, "teamName": _myTeam.displayName, "channelName": _myChannel.displayName ,"messageDateTime": _reply.createdDateTime ,"messageBy": _reply.from.user.displayName, "messageContent": _reply.body.content})
                          setTimeout(() => resolve("done!"), 1000);
                        }
                        ///
                      })                    
                    }
                  } catch (error) {  
                    console.log('unable to get channels', error);  
                  }
                })
                ////
              } catch (messageError) {  
                console.log('unable to get Messages', messageError);  
              }
            })
          }
          catch(myChannelsError)
          {
            console.log('Unable to get Channel', myChannelsError);   
          }
        })
      } catch (error) {  
        console.log('Unable to get teams', error);  
      }

    });    

    let result = await promise; // wait until the promise resolves (*)

    console.log("Result...");
    console.log(result);
    ////    

    ///
    var _finalResult = [];

    _resultMessages.map(async (message: any) => {
          var _replyArray = [];
      _resultReplies.map(async (reply: any) => {
          if(message.messageId == reply.replyToId)
          {
            //_replyArray.push({"teamId": reply.teamId, "channelId": reply.channelId, "messageId": reply.messageId, "replyId": reply.replyId, "replyToId": reply.replyToId, "teamName": reply.teamName, "channelName": reply.channelName ,"messageDateTime": reply.messageDateTime, "messageBy": reply.messageBy, "messageContent": reply.messageContent})
            _replyArray.push({
              // "teamId": reply.teamId, 
              // "channelId": reply.channelId, 
              // "messageId": reply.messageId, 
              // "replyId": reply.replyId, 
              // "replyToId": reply.replyToId, 
              //"teamName": reply.teamName, 
              //"channelName": reply.channelName ,
              "messageDateTime": reply.messageDateTime, 
              "messageBy": reply.messageBy, 
              "messageContent": reply.messageContent})            
          }
      });

      _finalResult.push({
        // "teamId": message.teamId, 
        // "channelId": message.channelId, 
        // "messageId": message.messageId, 
        "Team": message.teamName ,
        "Channel": message.channelName ,
        "DateTime": message.messageDateTime, 
        "MessageBy": message.messageBy, 
        "Text": message.messageContent, 
        "Reply": _replyArray})
    })
    console.log("Final Result...");
    console.log(_finalResult);

    this.setState({
      teamExportPackage: _finalResult
    });
    ///
}

handleDateChangeFrom = (date) => {
  console.log(date);
  this.setState({ fromDate: date });
};

handleDateChangeTo = (date) => {
  console.log(date);
  this.setState({ toDate: date });
};

generatePdf = () => {
  const element = document.getElementById('pdf-content');
  const opt = {
    margin: 10,
    filename: 'document.pdf',
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
  };

  // Generate PDF from the element and download it directly
  html2pdf(element, opt);
};


  public render(){
    let _index = 1;

    return (
      <div>
        <div>
            <div>
              <div style={{width:"100%", display: "inline-flex"}}>
                  <div style={{width:"20%", padding: "10px"}}>
                    <Dropdown
                        placeholder="Select an option"
                        label="Select Team"
                        options={this.state.allTeams}
                        onChanged={ this._onTeamsDropDownChanged.bind(this) }
                    />
                  </div>
                  <div style={{width:"20%", padding: "10px"}}>
                    <Dropdown
                        placeholder="Select an option"
                        label="Select Channel"
                        options={this.state.allChannels}
                        onChanged={ this._onChannelDropDownChanged.bind(this) }
                    />
                  </div>                
                  <div style={{width:"20%", padding: "10px"}}>
                  <DatePicker
                    label="Select from date"
                    value={this.state.fromDate}
                    onSelectDate={this.handleDateChangeFrom}
                  />
                  </div>                                  
                  <div style={{width:"20%", padding: "10px"}}>
                    <DatePicker
                      label="Select to date"
                      value={this.state.toDate}
                      onSelectDate={this.handleDateChangeTo}
                    />
                  </div>
                  <div style={{width:"20%", padding: "10px"}}>
                    <Dropdown
                          defaultSelectedKey={this.state.defaultSelectedKey}
                          selectedKey={this.state.defaultSelectedKey}
                          placeholder="Select an option"
                          label="Select Top Load"
                          options={this.state.topload}
                          onChanged={ this._onTopLoadDropDownChanged.bind(this) }
                      />
                  </div>                               
                  <div style={{width:"30%", padding: "10px", paddingTop: "37px"}}>
                    <PrimaryButton text="Get CSV" onClick={this.csvButtonClicked} />&nbsp;
                    <PrimaryButton text="Get PDF" onClick={this.pdfButtonClicked} />
                  </div>                  
                </div>
                <div style={{display: this.state.isCsvVisible}}>
                  <DetailsListDocumentsExample data={this.state._allItems}></DetailsListDocumentsExample>
                </div>
                <div style={{display: this.state.isPdfVisible}}>
                  <div>
                    <PrimaryButton text="Generate PDF" onClick={() => this.generatePdf()} />                 
                    <div id="pdf-content">
                          <div>
                            <div style={{maxWidth: "67%"}}>

                              <ul style={{listStyleType: "none", width: "90%"}} className="item">
                              {
                                  this.state.teamExportPackage.map((item, key) => (
                                    (item.ReplyID == 0 ? 
                                      <li key={item.ID}>
                                          <div style={{paddingLeft: "0px"}}>
                                              <div><b>{_index++}. {item.Sender} | {item.DateTime}</b></div>
                                              <div style={{textAlign: "justify", paddingLeft: "14px"}}>{item.Message}</div>
                                          </div>
                                      </li>
                                      :
                                      <li key={item.ID}>
                                          <div style={{paddingLeft: "25px", fontSize: "small"}}>
                                              <div><b>{item.Sender} | {item.DateTime}</b></div>
                                              <div style={{textAlign: "justify"}}>{item.Message}</div>
                                          </div>
                                      </li>
                                    )    
                                  ))
                              }
                              </ul>
                            </div>
                          </div>
                    </div>
                </div>

                </div>                
            </div>  
        </div>
      </div>
    );
  }
}
