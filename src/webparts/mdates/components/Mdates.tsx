import * as React from 'react';
import styles from './Mdates.module.scss';
import { IMdatesProps } from './IMdatesProps';
import { IMdatesState } from './IMdatesState';
import { DateRange } from 'react-date-range';
import { autobind } from 'office-ui-fabric-react/lib/utilities';
import 'react-date-range/dist/styles.css';
import 'react-date-range/dist/theme/default.css';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton } from 'office-ui-fabric-react';
import { escape } from '@microsoft/sp-lodash-subset';
import { Lists } from '@pnp/sp/lists';


export default class Mdates extends React.Component<IMdatesProps, IMdatesState, {}> {
  constructor(props: IMdatesProps, state: IMdatesState) {
    super(props);
    // sp.setup({ spfxcontext: this.props.context });
    this.state = ({ startDate: new Date(), endDate: null, key: 'selection' })
    this.getValuesFromSP();
  }
  public render(): React.ReactElement<IMdatesProps> {
    let state = [{ startDate: this.state.startDate, endDate: this.state.endDate, key: this.state.key }]

    return (
      <div className={styles.mdates}>
        <DateRange
          editableDateInputs={true}
          // onChange={item => this.setState({ endDate: this.selection["endDate"], startDate: this.selection["startDate"] })}
          moveRangeOnFirstSelection={false}
          ranges={state}
        />
        <br>
          <PrimaryButton text="Save" onClick={this._SaveIntoSP} />
        </br>
      </div>
    );
  }


  private async getValuesFromSP() {
    const item: any = await sp.web.lists.getByTitle("DateRangeList").items.getById(1).get();
    this.setState({ endDate: item.DateFrom, startDate: item.DateTo })

  }
  @autobind
  private async _SaveIntoSP() {
    let list = sp.web.lists.getByTitle("DateRangeList");
    const i = await list.items.getById(1).update({
      DateFrom: this.state.startDate,
      DateTo: this.state.endDate
    });
  }
}