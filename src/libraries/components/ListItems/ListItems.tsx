import * as React from 'react';
import { IListItems } from "./IListItems";
import { ListView, IViewField, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";

export default function ListItems (props: IListItems) {
    
    return(
        <ListView
            items={[]}
            // viewFields={viewFields}
            // groupByFields={groupByFields}
            // compact={true}
        /> 
    );
}