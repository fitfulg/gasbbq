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

const LANGUAGES = [
    {
        code: 'en',
        name: 'English',
        menuName: 'Language',
        headers: ['Name', 'Confirmation', 'Food Preference', 'Drink Preference', 'Allergies', 'C-counter (do not edit)', 'D-counter (do not edit)']
    },
    {
        code: 'es',
        name: 'Castellano',
        menuName: 'Idioma',
        headers: ['Nombre', 'Confirmación', 'Preferencia de Comida', 'Preferencia de Bebida', 'Alergias', 'C-contador (no editar)', 'D-contador (no editar)']
    },
    {
        code: 'ca',
        name: 'Català',
        menuName: 'Idioma',
        headers: ['Nom', 'Confirmació', 'Preferència menjars', 'Preferència begudes', 'Al·lèrgies', 'C-counter (no editar)', 'D-counter (no editar)']
    }
];

// module.exports = { COLORS, COLUMN_CONFIG, LANGUAGES };