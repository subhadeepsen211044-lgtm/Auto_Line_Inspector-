
export type ZoneId = 'L1' | 'L2' | 'L3' | 'L4' | 'R0' | 'R1' | 'R2' | 'R3' | 'R4';

export interface ManpowerEntry {
  name: string;
  ein: string;
}

export interface ZoneResponse {
  zoneId: ZoneId;
  involved: 'YES' | 'NO' | 'PENDING';
  manpower: ManpowerEntry[];
  reason?: string;
  actionTaken?: string;
  timestamp?: string;
}

export interface ManagerAnalysis {
  man: string;
  machine: string;
  material: string;
  method: string;
  rootCauseSummary: string;
  status: 'OPEN' | 'CLOSED';
  closedAt?: string;
}

export interface Defect {
  id: string;
  frameNo: string;
  modelName: string;
  defectName: string;
  description: string;
  photoUrl?: string;
  targetedZones: ZoneId[];
  timestamp: string;
  responses: Record<string, ZoneResponse>;
  managerAnalysis?: ManagerAnalysis;
}

export interface TsprEntry {
  slot: string;
  total: number;
  ng: number;
}

export type AppRole = 'INSPECTOR' | 'GROUP_LEADER' | 'MANAGER';

export interface UserSession {
  role: AppRole;
  zoneId?: ZoneId; // Only for Group Leaders
  name: string;
}