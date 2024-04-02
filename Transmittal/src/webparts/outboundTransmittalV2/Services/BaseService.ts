import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';

export class BaseService {
    private _sp: any;

    constructor(context: WebPartContext) {
        this._sp = SPFI;
    }

    public getLogException(Details: any): void {
        console.log(Details);
    }



}


