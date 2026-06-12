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

// Note: hair salon / barber categories are intentionally excluded — they are
// covered by EXCLUDED_BUSINESS_TYPES and would produce zero leads.
export const SEARCHES = [
  { category: "plumber",           towns: ESSEX_TOWNS },
  { category: "electrician",       towns: ESSEX_TOWNS },
  { category: "cafe",              towns: ESSEX_TOWNS },
  { category: "cleaning service",  towns: ESSEX_TOWNS },
  { category: "painter decorator", towns: ESSEX_TOWNS },
  { category: "gardener",          towns: ESSEX_TOWNS },
  { category: "handyman",          towns: ESSEX_TOWNS },
  { category: "plasterer",         towns: ESSEX_TOWNS },
  { category: "tiler",             towns: ESSEX_TOWNS },
  { category: "roofer",            towns: ESSEX_TOWNS },
  { category: "carpenter",         towns: ESSEX_TOWNS },
  { category: "locksmith",         towns: ESSEX_TOWNS },
  { category: "mobile mechanic",   towns: ESSEX_TOWNS },
];

export const MAX_RESULTS_PER_SEARCH = 20;

export const DELAYS = {
  betweenActions:  [800,  2000],
  betweenListings: [1500, 3500],
  betweenSearches: [4000, 8000],
};
