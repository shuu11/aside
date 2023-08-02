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

export class Sheet {
  #ss: TypeApp = SpreadsheetApp.getActiveSpreadsheet();

  constructor(public name: string) {
    this.name = name;
  }

  getSheet(): TypeSheet | null {
    return this.#ss.getSheetByName(this.name);
  }

  getValue(row: number, column: number): any {
    return this.getSheet()?.getRange(row, column)?.getValue();
  }

  getValues(row: number, column: number): any[][] | undefined {
    return this.getSheet()?.getRange(row, column)?.getValues();
  }

  getLastRow(): number | undefined {
    return this.getSheet()?.getLastRow();
  }

  getLastColumn(): number | undefined {
    return this.getSheet()?.getLastColumn();
  }

  setValue(row: number, column: number, value: any): void {
    this.getSheet()?.getRange(row, column)?.setValue(value);
  }
}
