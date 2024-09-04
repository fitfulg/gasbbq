const ColorUtils = {
    darkGray: () => '#4d4d4d',
    lightGray: () => '#d9d9d9',
    white: () => '#ffffff',
    lightYellow: () => '#ffffe6',
    lightBlue: () => '#e6f2ff',
};

const HEADERS_CONFIG = [
    'Nom',
    'Confirmació',
    'Preferència menjars',
    'Preferència begudes',
    'Al·lèrgies',
    'C-counter (no editar)',
    'D-counter (no editar)'
];

module.exports = { ColorUtils, HEADERS_CONFIG };