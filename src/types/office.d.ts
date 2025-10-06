/**
 * Additional type declarations for Office.js
 * These supplement the @types/office-js package
 */

// Extend Excel namespace with missing types
declare namespace Excel {
  interface RequestContext {
    workbook: Workbook;
    sync(): Promise<void>;
  }

  interface Workbook {
    name: string;
    worksheets: WorksheetCollection;
    getSelectedRange(): Range;
  }

  interface WorksheetCollection {
    items: Worksheet[];
    load(propertyNames?: string | string[]): void;
    getItem(name: string): Worksheet;
    getActiveWorksheet(): Worksheet;
  }

  interface Worksheet {
    name: string;
    getRange(address: string): Range;
    getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number): Range;
    getUsedRange(): Range;
    load(propertyNames?: string | string[]): void;
  }

  interface Range {
    address: string;
    formulas: any[][];
    values: any[][];
    worksheet: Worksheet;
    rowCount: number;
    columnCount: number;
    format: RangeFormat;
    select(): void;
    insert(shift: InsertShiftDirection): void;
    delete(shift: DeleteShiftDirection): void;
    load(propertyNames?: string | string[]): void;
  }

  interface RangeFormat {
    fill: RangeFill;
  }

  interface RangeFill {
    color: string;
    clear(): void;
  }

  enum InsertShiftDirection {
    down = "Down",
    right = "Right"
  }

  enum DeleteShiftDirection {
    up = "Up",
    left = "Left"
  }

  function run<T>(
    callback: (context: RequestContext) => Promise<T>
  ): Promise<T>;
}

declare namespace Office {
  function onReady(callback?: (info: { host: HostType; platform: PlatformType }) => void): Promise<{ host: HostType; platform: PlatformType }>;

  enum HostType {
    Word = "Word",
    Excel = "Excel",
    PowerPoint = "PowerPoint",
    Outlook = "Outlook",
    OneNote = "OneNote",
    Project = "Project",
    Access = "Access"
  }

  enum PlatformType {
    PC = "PC",
    OfficeOnline = "OfficeOnline",
    Mac = "Mac",
    iOS = "iOS",
    Android = "Android",
    Universal = "Universal"
  }

  namespace AddinCommands {
    interface Event {
      completed(options?: any): void;
    }
  }

  interface NotificationMessageDetails {
    type: any;
    message: string;
    icon: string;
    persistent: boolean;
  }

  namespace MailboxEnums {
    enum ItemNotificationMessageType {
      InformationalMessage = "InformationalMessage",
      ErrorMessage = "ErrorMessage",
      InsightMessage = "InsightMessage"
    }
  }

  const context: {
    mailbox?: {
      item?: {
        notificationMessages?: {
          replaceAsync(key: string, message: NotificationMessageDetails): void;
        };
      };
    };
  };
}
