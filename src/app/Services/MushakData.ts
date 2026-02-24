export interface MushakData {
  taxpayer: {
    bin: string;
    name: string;
    address: string;
    businessNature: string;
    activity: string;
  };
  returnSubmission: {
    period: string;
    type: string;
    hasActivity: boolean;
    date: string;
  };
  notes: {
    [key: string]: any;
  };
}