/**
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
import { TypeApp, TypeSheet } from '../types/types';

export type activeSheet = {
  getSheet(): TypeSheet;
  getSheetName(): string;
  getValue(...args: any): any;
  getValues(row: number, column: number): any[][];
  getLastRow(): number;
  getLastColumn(): number;
  setValue(row: number, column: number, value: any): void;
};

export class ActiveSheet implements activeSheet {
  #ss: TypeApp = SpreadsheetApp.getActiveSpreadsheet();

  getSheet(): TypeSheet {
    return this.#ss.getActiveSheet();
  }

  getSheetName(): string {
    return this.#ss.getActiveSheet().getName();
  }

  getValue(...args: any): any {
    let ret;

    switch (args.length) {
      case 2:
        const row = args[0];
        const column = args[1];

        ret = this.#ss.getActiveSheet().getRange(row, column).getValue();
        break;

      case 1:
        if (typeof args[0] === 'string') {
          const a1Notation = args[0];
          ret = this.#ss.getActiveSheet().getRange(a1Notation).getValue();
        } else {
          throw new Error('Invalid arguments.');
        }
        break;

      default:
        throw new Error('Invalid arguments.');
    }

    return ret;
  }

  getValues(row: number, column: number): any[][] {
    return this.#ss.getActiveSheet().getRange(row, column)?.getValues();
  }

  getLastRow(): number {
    return this.#ss.getActiveSheet().getLastRow();
  }

  getLastColumn(): number {
    return this.#ss.getActiveSheet().getLastColumn();
  }

  setValue(row: number, column: number, value: any): void {
    this.#ss.getActiveSheet().getRange(row, column).setValue(value);
  }
}
