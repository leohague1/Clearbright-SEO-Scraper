// Substring-matched against business name + GBP category (case-insensitive).
// Add new terms here to exclude additional business types.
export const EXCLUDED_BUSINESS_TYPES = [
  "barber",
  "barbers",
  "barber shop",
  "barbershop",
  "hair salon",
  "hairdresser",
  "hairdressing",
  "hair & beauty",
  "hair and beauty",
  "hair stylist",
  "hair studio",
  "hair cutting",
  "haircut",
  "grooming salon",
];

// All Essex towns — add new entries here to extend coverage across all search categories.
const ESSEX_TOWNS = [
  "Colchester", "Chelmsford", "Basildon", "Southend-on-Sea", "Braintree",
  "Brentwood", "Harlow", "Witham", "Maldon", "Epping", "Loughton",
  "Billericay", "Rayleigh", "Wickford", "Canvey Island", "Clacton-on-Sea",
  "Frinton-on-Sea", "Halstead", "Saffron Walden", "Great Dunmow",
  "Burnham-on-Crouch", "Hockley", "Rochford", "Stanford-le-Hope", "Grays",
  "Tilbury", "Corringham", "South Woodham Ferrers", "Tiptree", "Coggeshall",
  "Mersea Island", "Wivenhoe",
];

export const SEARCHES = [
  { category: "accountant", towns: ["Chelmsford", "Colchester", "Southend-on-Sea", "Basildon", "Braintree", "Harlow", "Brentwood", "Grays", "Thurrock", "Wickford", "Rayleigh", "Billericay", "Witham", "Clacton-on-Sea", "Maldon", "Saffron Walden", "Halstead", "Loughton", "Epping", "Waltham Abbey", "Romford", "Ilford", "Hornchurch", "Upminster", "Leigh-on-Sea", "Westcliff-on-Sea", "Canvey Island"] },
  { category: "bookkeeper", towns: ["Chelmsford", "Colchester", "Southend-on-Sea", "Basildon", "Braintree", "Harlow", "Brentwood", "Grays", "Thurrock", "Wickford", "Rayleigh", "Billericay", "Witham", "Clacton-on-Sea", "Maldon", "Saffron Walden", "Halstead", "Loughton", "Epping", "Waltham Abbey", "Romford", "Ilford", "Hornchurch", "Upminster", "Leigh-on-Sea", "Westcliff-on-Sea", "Canvey Island"] },
  { category: "solicitor", towns: ["Chelmsford", "Colchester", "Southend-on-Sea", "Basildon", "Braintree", "Harlow", "Brentwood", "Grays", "Thurrock", "Wickford", "Rayleigh", "Billericay", "Witham", "Clacton-on-Sea", "Maldon", "Saffron Walden", "Halstead", "Loughton", "Epping", "Waltham Abbey", "Romford", "Ilford", "Hornchurch", "Upminster", "Leigh-on-Sea", "Westcliff-on-Sea", "Canvey Island"] },
  { category: "dentist", towns: ["Chelmsford", "Colchester", "Southend-on-Sea", "Basildon", "Braintree", "Harlow", "Brentwood", "Grays", "Thurrock", "Wickford", "Rayleigh", "Billericay", "Witham", "Clacton-on-Sea", "Maldon", "Saffron Walden", "Halstead", "Loughton", "Epping", "Waltham Abbey", "Romford", "Ilford", "Hornchurch", "Upminster", "Leigh-on-Sea", "Westcliff-on-Sea", "Canvey Island"] },
  { category: "estate agent", towns: ["Chelmsford", "Colchester", "Southend-on-Sea", "Basildon", "Braintree", "Harlow", "Brentwood", "Grays", "Thurrock", "Wickford", "Rayleigh", "Billericay", "Witham", "Clacton-on-Sea", "Maldon", "Saffron Walden", "Halstead", "Loughton", "Epping", "Waltham Abbey", "Romford", "Ilford", "Hornchurch", "Upminster", "Leigh-on-Sea", "Westcliff-on-Sea", "Canvey Island"] },
  { category: "gym", towns: ["Chelmsford", "Colchester", "Southend-on-Sea", "Basildon", "Braintree", "Harlow", "Brentwood", "Grays", "Thurrock", "Wickford", "Rayleigh", "Billericay", "Witham", "Clacton-on-Sea", "Maldon", "Saffron Walden", "Halstead", "Loughton", "Epping", "Waltham Abbey", "Romford", "Ilford", "Hornchurch", "Upminster", "Leigh-on-Sea", "Westcliff-on-Sea", "Canvey Island"] },
  { category: "physiotherapist", towns: ["Chelmsford", "Colchester", "Southend-on-Sea", "Basildon", "Braintree", "Harlow", "Brentwood", "Grays", "Thurrock", "Wickford", "Rayleigh", "Billericay", "Witham", "Clacton-on-Sea", "Maldon", "Saffron Walden", "Halstead", "Loughton", "Epping", "Waltham Abbey", "Romford", "Ilford", "Hornchurch", "Upminster", "Leigh-on-Sea", "Westcliff-on-Sea", "Canvey Island"] },
];

export const MAX_RESULTS_PER_SEARCH = 20;

export const DELAYS = {
  betweenActions:  [800,  2000],
  betweenListings: [1500, 3500],
  betweenSearches: [4000, 8000],
};
