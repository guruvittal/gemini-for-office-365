/**
 * Host Adapter Factory for Microsoft 365 (Word, PowerPoint, Excel)
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { WordAdapter } from './WordAdapter.js';
import { PPTAdapter } from './PPTAdapter.js';
import { ExcelAdapter } from './ExcelAdapter.js';

export class HostAdapterFactory {
  static getAdapter(info = null) {
    // 1. Direct host from Office.onReady(info)
    let host = info?.host;

    // 2. Diagnostics host
    if (!host && typeof Office !== 'undefined' && Office.context?.diagnostics?.host) {
      host = Office.context.diagnostics.host;
    }

    // 3. Fallback context host
    if (!host && typeof Office !== 'undefined' && Office.context?.host) {
      host = Office.context.host;
    }

    // 4. Runtime namespace inspection
    if (!host) {
      if (typeof PowerPoint !== 'undefined') {
        host = Office.HostType?.PowerPoint || "PowerPoint";
      } else if (typeof Excel !== 'undefined') {
        host = Office.HostType?.Excel || "Excel";
      } else if (typeof Word !== 'undefined') {
        host = Office.HostType?.Word || "Word";
      }
    }

    // 5. Requirements set inspection
    if (!host && typeof Office !== 'undefined' && Office.context?.requirements) {
      try {
        if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
          host = Office.HostType?.PowerPoint || "PowerPoint";
        } else if (Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
          host = Office.HostType?.Excel || "Excel";
        } else if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {
          host = Office.HostType?.Word || "Word";
        }
      } catch (_) {}
    }

    if (host === Office?.HostType?.PowerPoint || host === "PowerPoint") {
      return new PPTAdapter();
    }
    if (host === Office?.HostType?.Excel || host === "Excel") {
      return new ExcelAdapter();
    }
    if (host === Office?.HostType?.Word || host === "Word") {
      return new WordAdapter();
    }

    return new WordAdapter();
  }
}
