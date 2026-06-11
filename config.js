export const SEARCHES = [
  { category: "plumber", towns: ["Colchester", "Chelmsford", "Basildon", "Southend-on-Sea", "Braintree"] },
  { category: "electrician", towns: ["Colchester", "Chelmsford", "Basildon", "Southend-on-Sea"] },
  { category: "hair salon", towns: ["Colchester", "Chelmsford", "Basildon", "Southend-on-Sea"] },
  { category: "barber", towns: ["Colchester", "Chelmsford", "Basildon", "Southend-on-Sea"] },
  { category: "cafe", towns: ["Colchester", "Chelmsford"] },
  { category: "cleaning service", towns: ["Colchester", "Chelmsford", "Basildon"] },
  { category: "painter decorator", towns: ["Colchester", "Chelmsford", "Basildon"] },
];

export const MAX_RESULTS_PER_SEARCH = 20;

export const DELAYS = {
  betweenActions: [800, 2000],
  betweenListings: [1500, 3500],
  betweenSearches: [4000, 8000],
  betweenEmailSearches: [2000, 4000],
};
