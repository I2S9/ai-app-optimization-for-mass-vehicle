export const NAV_ITEMS = [
  { id: 'database', label: 'Database', path: '/database' },
  { id: 'synthesis', label: 'Synthesis', path: '/synthesis' },
  {
    id: 'cdc',
    label: 'CDC',
    path: '/cdc',
    children: [
      { id: 'cdc-mns', label: 'MNS', path: '/cdc/mns' },
      { id: 'cdc-options-sp2', label: 'Options SP2', path: '/cdc/options-sp2' },
      { id: 'cdc-output', label: 'Output for CDC', path: '/cdc/output' },
    ],
  },
  {
    id: 'cdg',
    label: 'CDG - Center of gravity',
    path: '/cdg',
    children: [
      { id: 'cdg-input-1', label: 'INPUT 1', path: '/cdg/input-1' },
      { id: 'cdg-input-2', label: 'INPUT 2', path: '/cdg/input-2' },
      { id: 'cdg-output', label: 'OUTPUT', path: '/cdg/output' },
    ],
  },
  { id: 'waterline', label: 'Waterline', path: '/waterline' },
  { id: 'weight-tax', label: 'Weight Tax', path: '/weight-tax' },
  { id: 'target-status-forecast', label: 'Target-status-forecast', path: '/target-status-forecast' },
  { id: 'graph-ptf-limits', label: 'Graph PTF Limits', path: '/graph-ptf-limits' },
  { id: 'portfolio', label: 'Portfolio', path: '/portfolio' },
  { id: 'split', label: 'Split', path: '/split' },
];

// Flattened list of every navigable route (top-level items + nested children),
// used for route validation and current-route lookup.
export const NAV_ROUTES = NAV_ITEMS.flatMap((item) =>
  item.children ? [item, ...item.children] : [item]
);

export const DEFAULT_ROUTE = 'database';
