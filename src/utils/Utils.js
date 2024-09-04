const COLORS = {
    darkGray: () => '#4d4d4d',
    lightGray: () => '#d9d9d9',
    white: () => '#ffffff',
    lightYellow: () => '#ffffe6',
    lightBlue: () => '#e6f2ff',
};

const COLUMN_CONFIG = [
    { column: 'A', name: 'Nom', width: 150 },
    { column: 'B', name: 'Confirmació', width: 150 },
    { column: 'C', name: 'Preferència menjars', width: 300 },
    { column: 'D', name: 'Preferència begudes', width: 300 },
    { column: 'E', name: 'Al·lèrgies', width: 100 },
    { column: 'F', name: 'C-counter (no editar)', width: 200 },
    { column: 'G', name: 'D-counter (no editar)', width: 200 }
];


module.exports = { COLORS, COLUMN_CONFIG };