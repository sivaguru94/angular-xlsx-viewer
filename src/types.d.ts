declare module '@univerjs/presets' {
  export function createUniver(options: any): { univer: any; univerAPI: any };
  export const LocaleType: { EN_US: string; ZH_CN: string };
  export function merge(...args: any[]): any;
}

declare module '@univerjs/preset-sheets-core' {
  export function UniverSheetsCorePreset(config?: any): any;
}

declare module '@univerjs/preset-sheets-core/locales/en-US' {
  const locale: any;
  export default locale;
}

declare module '@zwight/luckyexcel' {
  const LuckyExcel: {
    transformExcelToUniver(
      file: File,
      onSuccess: (data: any) => void,
      onError: (error: any) => void
    ): void;
  };
  export default LuckyExcel;
}
