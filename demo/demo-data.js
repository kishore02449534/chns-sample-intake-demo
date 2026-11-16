const DEMO_REFERENCE = {
    farms: [
        { farmName: "Synthetic Farm Alpha", location: "Testville, UT" },
        { farmName: "Synthetic Farm Beta", location: "Sampletown, ID" },
        { farmName: "Synthetic Farm Gamma", location: "Fakeburg, WY" },
    ],
    sampleTypeByCategory: {
        meat: "Animal",
        dairy: "Animal",
        eggs: "Animal",
        seafood: "Animal",
        vegetables: "Plant",
        fruits: "Plant",
        grain: "Plant",
        nutsseedslegumes: "Plant",
        foragefeed: "Feed",
        soil: "Other",
    },
    foods: [
        {
            foodName: "Cow (Beef)",
            breeds: ["Angus", "Red Angus", "Unknown"],
            items: ["Angus|Ribeye", "Angus|Ground", "Red Angus|Ribeye", "Unknown|Ribeye"],
        },
        { foodName: "Chicken", breeds: ["Broiler"], items: ["Broiler|Thigh"] },
        { foodName: "Kale", breeds: ["Red Russian", "Unknown"], items: ["Red Russian|Leaves", "Unknown|Null"] },
        { foodName: "Alfalfa", breeds: ["Common"], items: ["Common|Whole"] },
        { foodName: "Blueberries", breeds: ["Duke"], items: ["Duke|Whole"] },
        { foodName: "Egg (Chicken)", breeds: ["Unknown"], items: ["Unknown|Whole"] },
    ],
    speciesByCategoryWord: {
        "eggs:chicken": "Egg (Chicken)",
        ":egg": "Egg (Chicken)",
        ":blueberry": "Blueberries",
    },
    cuts: ["Ribeye", "Ground", "Thigh", "Leaves", "Whole", "Null"],
    cutBySheetWord: { wholeegg: "Whole", egg: "Whole" },
    practices: ["Grass-Fed", "Regenerative", "Conventional", "Pasture-Raised", "Organic"],
    breedWhenBlank: "Unknown",
    cutWhenBlank: "Null",
};

const DEMO_EXISTING_ROWS = [
    ["BFI-9001", "SYN Beef Ribeye 1", "10/02/2026", "Meat", "Beef", "Ribeye", "Angus", "Grass-Fed", "Synthetic Farm Alpha"],
    ["BFI-9002", "SYN Beef Ribeye 2", "10/02/2026", "Meat", "Beef", "Ribeye", "Angus", "Grass-Fed", "Synthetic Farm Alpha"],
    ["BFI-9003", "SYN Kale 1", "10/09/2026", "Vegetables", "Kale", "Leaves", "Red Russian", "Organic", "Synthetic Farm Beta"],
    ["BFI-9004", "SYN Kale 2", "10/09/2026", "Vegetables", "Kale", "Leaves", "Red Russian", "", "Synthetic Farm Beta"],
    ["BFI-9005", "SYN Alfalfa 1", "10/21/2026", "Forage/Feed", "Alfalfa", "Whole", "Common", "Conventional", "Synthetic Farm Gamma"],
    ["BFI-9006", "SYN Eggs 1", "10/28/2026", "Eggs", "Chicken", "Whole Egg", "", "Pasture-Raised", "Synthetic Farm Gamma"],
];

const DEMO_GOOD_EXAMPLES = [
    { tubeLabel: "SYN Beef Ribeye", category: "Meat", species: "Beef", cut: "Ribeye", breed: "Angus", practice: "Grass-Fed", farmName: "Synthetic Farm Alpha" },
    { tubeLabel: "SYN Kale", category: "Vegetables", species: "Kale", cut: "Leaves", breed: "Red Russian", practice: "Organic", farmName: "Synthetic Farm Beta" },
    { tubeLabel: "SYN Eggs", category: "Eggs", species: "Chicken", cut: "Whole Egg", breed: "", practice: "Pasture-Raised", farmName: "Synthetic Farm Gamma" },
    { tubeLabel: "SYN Blueberries", category: "Fruits", species: "Blueberries", cut: "Whole", breed: "Duke", practice: "Organic, Regenerative", farmName: "Synthetic Farm Beta" },
    { tubeLabel: "SYN Chicken Thigh", category: "Meat", species: "Chicken", cut: "Thigh", breed: "Broiler", practice: "Pasture-Raised", farmName: "Synthetic Farm Gamma" },
    { tubeLabel: "SYN Ground Beef", category: "Meat", species: "Beef", cut: "Ground", breed: "Angus", practice: "Regenerative", farmName: "Synthetic Farm Alpha" },
];

const DEMO_PROBLEM_EXAMPLES = [
    { tubeLabel: "SYN Beef Ribeye", category: "Meat", species: "Beef", cut: "Ribeye", breed: "Angus", practice: "Grass-Fed", farmName: "Green Acre Farm" },
    { tubeLabel: "SYN Llama", category: "Meat", species: "Llama", cut: "Ground", breed: "", practice: "", farmName: "Synthetic Farm Alpha" },
    { tubeLabel: "SYN Beef Brisket", category: "Meat", species: "Beef", cut: "Brisket", breed: "Angus", practice: "", farmName: "Synthetic Farm Alpha" },
    { tubeLabel: "SYN Kale", category: "Vegetables", species: "Kale", cut: "Leaves", breed: "Red Russian", practice: "Biodynamic", farmName: "Synthetic Farm Beta" },
    { tubeLabel: "SYN Kibble", category: "Pet food", species: "Chicken", cut: "Thigh", breed: "Broiler", practice: "", farmName: "Synthetic Farm Gamma" },
    { tubeLabel: "SYN Kale", category: "Vegetables", species: "Kale", cut: "Leaves", breed: "Red Russian", practice: "", farmName: "Synthetic Farm Beta", receivedDate: "next week" },
];
