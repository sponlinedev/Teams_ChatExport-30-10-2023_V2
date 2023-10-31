import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Announced } from '@fluentui/react/lib/Announced';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { TooltipHost } from '@fluentui/react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import csvDownload from 'json-to-csv-export';

const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px',
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden',
      },
    },
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '16px',
    maxWidth: '16px',
  },
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap',
  },
  exampleToggle: {
    display: 'inline-block',
    marginBottom: '10px',
    marginRight: '30px',
  },
  selectionDetails: {
    marginBottom: '20px',
  },
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '300px',
  },
};


export interface IDetailsListDocumentsExampleState {
  columns: IColumn[];
  items: IDocument[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
  announcedMessage?: string;
}

export interface IDocument {
  key: string;
  name: string;
  value: string;
  iconName: string;
  fileType: string;
  modifiedBy: string;
  dateModified: string;
  dateModifiedValue: number;
  fileSize: string;
  fileSizeRaw: number;
}

export class DetailsListDocumentsExample extends React.Component<any, any> {
  private _selection: Selection;
  private _allItems: IDocument[];

  constructor(props: {}) {
    super(props);

    const columns: IColumn[] = [
    {
        key: 'column1',
        name: 'ID',
        fieldName: 'ID',
        minWidth: 19,
        maxWidth: 20,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        isPadded: true,
    },
    {
        key: 'column2',
        name: 'ReplyID',
        fieldName: 'ReplyID',
        minWidth: 29,
        maxWidth: 30,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        isPadded: true,
    },    
      {
        key: 'column3',
        name: 'Message',
        fieldName: 'Message',
        minWidth: 149,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column4',
        name: 'Sender',
        fieldName: 'Sender',
        minWidth: 79,
        maxWidth: 80,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column5',
        name: 'DateTime',
        fieldName: 'DateTime',
        minWidth: 89,
        maxWidth: 95,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column6',
        name: 'Type',
        fieldName: 'MessageType',
        minWidth: 79,
        maxWidth: 80,
        isResizable: true,
        isCollapsible: false,
        data: 'string',
      },                 
    ];

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails(),
        });
      },
    });

    this.state = {
      items: this.props.data, // this._allItems,
      columns: columns,
      selectionDetails: "", // this._getSelectionDetails(),
      isModalSelection: true,
      isCompactMode: false,
      announcedMessage: undefined,
      selectedrows: []
    };
  }

  private downloadCSV() {
    var _selectedRows = [];

    this.state.selectedrows.map((item: any) => {
      _selectedRows.push({
        "ID": item["ID"],
        "ReplyID": item["ReplyID"],
        "MessageCSV": item["MessageCSV"],
        "Sender": item["Sender"],
        "DateTime": item["DateTime"],
        "MessageType": item["MessageType"]
      });
    })


    const dataToConvert = {
      data: _selectedRows, //this.state.selectedrows,
      filename: 'TeamChatExport',
      delimiter: ',',
      headers: ['ID', 'ReplyID', 'MessageCSV', 'Sender', 'DateTime', 'MessageType']
    }

    csvDownload(dataToConvert);
  }  


  public render() {
    const { columns, isCompactMode, items, selectionDetails, isModalSelection, announcedMessage } = this.state;
    return (
      <div>
        {this.props.data && this.props.data.length > 0  ?
          <div>
              <div>
                <PrimaryButton text="Export CSV" onClick={() => this.downloadCSV()} />
              </div>
              <div className={classNames.controlWrapper}>
                <DetailsList
                  items={this.props.data}
                  compact={isCompactMode}
                  columns={columns}
                  selectionMode={SelectionMode.multiple}
                  getKey={this._getKey}
                  setKey="multiple"
                  layoutMode={DetailsListLayoutMode.justified}
                  isHeaderVisible={true}
                  selection={this._selection}
                  selectionPreservedOnEmptyClick={true}
                  onItemInvoked={this._onItemInvoked}
                  enterModalSelectionOnTouch={true}
                  ariaLabelForSelectionColumn="Toggle selection"
                  ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                  checkButtonAriaLabel="select row"
                />
              </div>       
            </div>       
        :
          <></>
        }
      </div>
    );
  }

  public componentDidUpdate(previousProps: any, previousState: IDetailsListDocumentsExampleState) {
    if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
      this._selection.setAllSelected(false);
    }
  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }

  private _onItemInvoked(item: any): void {
    alert(`Item invoked: ${item.name}`);
  }

  private _getSelectionDetails() {
    const selectionCount = this._selection.getSelectedCount();

    this.setState({
      selectedrows: this._selection.getSelection()
    })
  }
}
