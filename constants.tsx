
import React from 'react';
import { ShieldAlert, Users, LayoutDashboard, Settings, ClipboardList } from 'lucide-react';
import { ZoneId } from './types';

export const ZONES: ZoneId[] = ['L1', 'L2', 'L3', 'L4', 'R0', 'R1', 'R2', 'R3', 'R4'];
export const LEFT_ZONES: ZoneId[] = ['L1', 'L2', 'L3', 'L4'];
export const RIGHT_ZONES: ZoneId[] = ['R0', 'R1', 'R2', 'R3', 'R4'];

export const ROLE_CONFIG = {
  INSPECTOR: {
    title: 'Final Inspector',
    icon: <ShieldAlert className="w-6 h-6" />,
    color: 'bg-blue-600',
  },
  GROUP_LEADER: {
    title: 'Group Leader',
    icon: <Users className="w-6 h-6" />,
    color: 'bg-orange-600',
  },
  MANAGER: {
    title: 'Line Manager',
    icon: <LayoutDashboard className="w-6 h-6" />,
    color: 'bg-indigo-600',
  }
};
