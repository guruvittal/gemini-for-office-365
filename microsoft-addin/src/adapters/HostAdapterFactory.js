/**
 * Host Adapter Factory for Microsoft 365 (Word, PowerPoint, Excel)
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { WordAdapter } from './WordAdapter.js';
import { PPTAdapter } from './PPTAdapter.js';
import { ExcelAdapter } from './ExcelAdapter.js';

export class HostAdapterFactory {
  static getAdapter() {
    if (typeof Office !== 'undefined' && Office.context && Office.context.host) {
      switch (Office.context.host) {
        case Office.HostType.Word:
          return new WordAdapter();
        case Office.HostType.PowerPoint:
          return new PPTAdapter();
        case Office.HostType.Excel:
          return new ExcelAdapter();
        default:
          return new WordAdapter();
      }
    }
    return new WordAdapter();
  }
}
