// Built-in baseline: everything the dashboard used to read from
// data/source-workbook.xlsx when the page opens (project values, forecast
// inputs, INFORMATION values, manpower, advanced settings incl. the Electricity
// table, Reference (LocationType & Density) table, GP% / Basket / Write-off
// lookups from 'Sales forecasting tools'!AB:AE, Fr Com% vs GP Share% table,
// category sales mix and signatories). It is the exact result of reading that
// workbook, so the dashboard opens with the same figures and rules without any
// Excel file in the repository. Edit values here to change the baseline.
export const BASELINE_DATA = {
  "meta": {
    "sourceName": "source-workbook.xlsx",
    "sourceLoaded": true
  },
  "project": {
    "locationArea": "Khilkhet Amtola, Hazibari, House - Kha-53, Khilkhet, Dhaka -1229",
    "googleMapLocation": "",
    "district": "",
    "division": "Dhaka",
    "pnp": "N",
    "frOwn": "FR",
    "sft": 1550,
    "density": "H",
    "incomeLevel": "A",
    "locationType": "Commercial Hub",
    "longFeet": 25,
    "projectedDailySales": 60000,
    "monthlySalesOverride": null,
    "monthlyRent": 61000,
    "advance": 1200000,
    "outboundTransport": 10000,
    "salesGivenBy": "Mr. Ersadul Haque",
    "openedBy": "Mr. Fakhrul Alamgir",
    "openedDesignation": "Head Of Business Expansion",
    "existingOutlets": 0,
    "includeExistingOutlets": true,
    "gpPercentOverride": null,
    "gpShareOverride": null
  },
  "forecast": {
    "marketNearby": "",
    "avgDepartmentalSales": 25000,
    "roadStatus": "M",
    "worshipCount": 2,
    "educationCount": 2,
    "bankOfficeCount": 2,
    "competitorAvgSales": 0,
    "publicTransit": "Y",
    "signboardVisibility": "H",
    "hotelRestaurantHospitalCount": 3
  },
  "information": {
    "manpowerAuto": true,
    "otherIncomeRate": 0.03,
    "cepValueOverride": null,
    "decorationCostOverride": null,
    "basketSizeOverride": null,
    "footfallOverride": null,
    "areaOutsideDhaka": "N"
  },
  "staff": [
    {
      "id": "om",
      "name": "OM-Parmanent",
      "group": "permanent",
      "quantity": 1,
      "salary": 23000
    },
    {
      "id": "icmo",
      "name": "ICMO-Parmanent",
      "group": "permanent",
      "quantity": 0,
      "salary": 15000
    },
    {
      "id": "duty",
      "name": "Duty officer",
      "group": "contractual",
      "quantity": 1,
      "salary": 12000
    },
    {
      "id": "cg",
      "name": "CG",
      "group": "contractual",
      "quantity": 2,
      "salary": 8000
    },
    {
      "id": "commodity",
      "name": "COMMODITY",
      "group": "contractual",
      "quantity": 0,
      "salary": 8000
    },
    {
      "id": "protein",
      "name": "PROTEIN",
      "group": "contractual",
      "quantity": 0,
      "salary": 10000
    },
    {
      "id": "perishables",
      "name": "PERISHABLES",
      "group": "contractual",
      "quantity": 0,
      "salary": 8000
    },
    {
      "id": "gml",
      "name": "GM & LIFESTYLE",
      "group": "contractual",
      "quantity": 0,
      "salary": 8000
    },
    {
      "id": "pos",
      "name": "POS",
      "group": "contractual",
      "quantity": 1,
      "salary": 9500
    },
    {
      "id": "porter",
      "name": "PORTER",
      "group": "contractual",
      "quantity": 0,
      "salary": 7500
    },
    {
      "id": "bsm",
      "name": "BSM",
      "group": "contractual",
      "quantity": 0,
      "salary": 11000
    },
    {
      "id": "bkstr",
      "name": "BK STR",
      "group": "contractual",
      "quantity": 0,
      "salary": 8000
    },
    {
      "id": "security",
      "name": "SECURITY",
      "group": "support",
      "quantity": 0,
      "salary": 11500
    },
    {
      "id": "cleaner",
      "name": "CLEANER",
      "group": "support",
      "quantity": 0,
      "salary": 7500
    }
  ],
  "advanced": {
    "stockPerSft": 1650,
    "stockFreeHoldingDays": 55,
    "gpAnnualStep": 0.002,
    "salesGrowthYear2": 0.12,
    "salesGrowthYear3": 0.1,
    "salesGrowthYear4": 0.1,
    "salesGrowthYear5": 0.08,
    "basketGrowth": 0.04,
    "staffEscalation": 0.08,
    "outletOpexInitial": 25000,
    "outletOpexRecurringMonthly": 5000,
    "outletOpexEscalation": 0.05,
    "consumptionRate": 0.0065,
    "consumptionRatePnp": 0.008,
    "productWastageRate": 0.0058,
    "productWastageFirstMonthPnpRate": 0.0023,
    "productWastageFirstMonthNonPnpRate": 0.0006,
    "electricityMonthly": 0,
    "maintenanceMonthly": 0,
    "securityCostMonthly": 0,
    "generatorMonthly": 0,
    "cleaningMonthly": 0,
    "cityChargeDhakaMonthly": 4500,
    "cityChargeOutsideDhakaMonthly": 6500,
    "membershipDiscountRate": 0.0025,
    "insuranceMonthly": 2500,
    "promotionalMonthly": 0,
    "iceMonthly": 0,
    "denominationRate": 0.0003,
    "creditCardRate": 0.003,
    "officeCostEscalation": 0.05,
    "conveyanceMonthly": 4000,
    "printingMonthly": 2500,
    "entertainmentMonthly": 1000,
    "stockWriteOffRate": 0.0048,
    "outletFinanceRate": 0.14,
    "outletDepreciablePortion": 0.7,
    "outletDepreciationMonths": 60,
    "transportEscalation": 0.05,
    "franchiseElectricityMonthlyOverride": null,
    "electricityTable": {
      "0 TO 800N": 18774.330577777775,
      "801 TO 1000N": 23712.098533333334,
      "1001 TO 1200N": 26357.078560000005,
      "1201 TO 1500N": 25954.704008163277,
      "1201 TO 1500Y": 49610.5858,
      "1501 TO 1800N": 27788.094278947363,
      "1501 TO 1800Y": 78564.1444,
      "1801 TO 2000N": 29175.871638709676,
      "1801 TO 2000Y": 75948.1788,
      "2001 TO 2500N": 38158.006250000006,
      "2501 TO 3000N": 49987.858912,
      "2501 TO 3000Y": 84756.352,
      "3001 TO 5000N": 48263.94507368421,
      "3001 TO 5000Y": 192027.42265000002,
      "5001 TO 8000N": 52838.538,
      "5001 TO 8000Y": 217743.5488,
      "2001 TO 2500Y": 70008
    },
    "franchiseElectricityEscalation": 0.03,
    "rentVatRate": 0.15,
    "franchiseMaintenanceMonthly": 2000,
    "franchiseGeneratorMonthly": 2000,
    "franchiseIceMonthly": 0,
    "franchiseServiceMonthly": 0,
    "franchiseFinanceRate": 0.14,
    "franchiseDepreciationMonths": 120,
    "rentEscalation": 0.1,
    "rentEscalationStartsYear": 4,
    "discountRate": 0.09,
    "securityDeposit": 0,
    "terminalRecovery": 3000000
  },
  "reference": {
    "locationDensityReference": {
      "sheetName": "Reference (LocationType & Dens ",
      "lastRow": 246,
      "rows": [
        {
          "locationType": "Commercial Hub",
          "density": "H",
          "division": "Barishal",
          "key": "Commercial HubHBarishal",
          "monthlySales": 4387198,
          "monthlyFootfall": 7620,
          "basket": 575.7477690288714,
          "profit": 187701.59062438027
        },
        {
          "locationType": "Commercial Hub",
          "density": "H",
          "division": "Chattogram",
          "key": "Commercial HubHChattogram",
          "monthlySales": 3002746.1428571427,
          "monthlyFootfall": 5077.714285714285,
          "basket": 591.3578381724061,
          "profit": 231129.94394007602
        },
        {
          "locationType": "Commercial Hub",
          "density": "H",
          "division": "Dhaka",
          "key": "Commercial HubHDhaka",
          "monthlySales": 3573803.4375,
          "monthlyFootfall": 6517.3125,
          "basket": 548.3553899709427,
          "profit": 257639.5005627723
        },
        {
          "locationType": "Commercial Hub",
          "density": "H",
          "division": "DhakaGBUD",
          "key": "Commercial HubHDhakaGBUD",
          "monthlySales": 1887336,
          "monthlyFootfall": 5767,
          "basket": 327.2647823825212,
          "profit": 89633.5986152569
        },
        {
          "locationType": "Commercial Hub",
          "density": "H",
          "division": "Mymensingh",
          "key": "Commercial HubHMymensingh",
          "monthlySales": 2873429.3333333335,
          "monthlyFootfall": 5657,
          "basket": 507.94225443403457,
          "profit": 204543.76496458054
        },
        {
          "locationType": "Commercial Hub",
          "density": "H",
          "division": "Rajshahi",
          "key": "Commercial HubHRajshahi",
          "monthlySales": 4299581.428571428,
          "monthlyFootfall": 7577,
          "basket": 567.4516864948433,
          "profit": 345932.6963153067
        },
        {
          "locationType": "Commercial Hub",
          "density": "H",
          "division": "Rangpur",
          "key": "Commercial HubHRangpur",
          "monthlySales": 2622345.3333333335,
          "monthlyFootfall": 4846.666666666667,
          "basket": 541.0616231086658,
          "profit": 68957.55813391386
        },
        {
          "locationType": "Commercial Hub",
          "density": "L",
          "division": "Barishal",
          "key": "Commercial HubLBarishal",
          "monthlySales": 774963.5,
          "monthlyFootfall": 2387.25,
          "basket": 324.6260341397005,
          "profit": -181.4932565524805
        },
        {
          "locationType": "Commercial Hub",
          "density": "L",
          "division": "Chattogram",
          "key": "Commercial HubLChattogram",
          "monthlySales": 1240578.875,
          "monthlyFootfall": 2583.625,
          "basket": 480.16986791813827,
          "profit": 65745.57195043395
        },
        {
          "locationType": "Commercial Hub",
          "density": "L",
          "division": "Dhaka",
          "key": "Commercial HubLDhaka",
          "monthlySales": 1247453.761904762,
          "monthlyFootfall": 3001.095238095238,
          "basket": 415.66616949367693,
          "profit": 57830.305659656944
        },
        {
          "locationType": "Commercial Hub",
          "density": "L",
          "division": "Khulna",
          "key": "Commercial HubLKhulna",
          "monthlySales": 1709327,
          "monthlyFootfall": 3937.75,
          "basket": 434.0872325566631,
          "profit": 114852.04689641047
        },
        {
          "locationType": "Commercial Hub",
          "density": "L",
          "division": "Mymensingh",
          "key": "Commercial HubLMymensingh",
          "monthlySales": 1033326.6,
          "monthlyFootfall": 2525,
          "basket": 409.2382574257426,
          "profit": 20861.937033661998
        },
        {
          "locationType": "Commercial Hub",
          "density": "L",
          "division": "Rajshahi",
          "key": "Commercial HubLRajshahi",
          "monthlySales": 1115810,
          "monthlyFootfall": 2670.1428571428573,
          "basket": 417.8840083462629,
          "profit": 40803.00531788128
        },
        {
          "locationType": "Commercial Hub",
          "density": "L",
          "division": "Rangpur",
          "key": "Commercial HubLRangpur",
          "monthlySales": 1187427.4285714286,
          "monthlyFootfall": 2968.285714285714,
          "basket": 400.0381172393879,
          "profit": 49906.51718178068
        },
        {
          "locationType": "Commercial Hub",
          "density": "L",
          "division": "Sylhet",
          "key": "Commercial HubLSylhet",
          "monthlySales": 743721,
          "monthlyFootfall": 1602,
          "basket": 464.24531835205994,
          "profit": -8258.145036551534
        },
        {
          "locationType": "Commercial Hub",
          "density": "M",
          "division": "Barishal",
          "key": "Commercial HubMBarishal",
          "monthlySales": 969423.1764705882,
          "monthlyFootfall": 2754.1176470588234,
          "basket": 351.99047415634345,
          "profit": 16589.619648938773
        },
        {
          "locationType": "Commercial Hub",
          "density": "M",
          "division": "Chattogram",
          "key": "Commercial HubMChattogram",
          "monthlySales": 1733135.0333333334,
          "monthlyFootfall": 3689,
          "basket": 469.81161109605137,
          "profit": 107748.46720804587
        },
        {
          "locationType": "Commercial Hub",
          "density": "M",
          "division": "Dhaka",
          "key": "Commercial HubMDhaka",
          "monthlySales": 1846852.0806451612,
          "monthlyFootfall": 4162.3387096774195,
          "basket": 443.7053804274117,
          "profit": 112043.44771317618
        },
        {
          "locationType": "Commercial Hub",
          "density": "M",
          "division": "DhakaGBUD",
          "key": "Commercial HubMDhakaGBUD",
          "monthlySales": 1069204,
          "monthlyFootfall": 3194,
          "basket": 334.7539135879775,
          "profit": -82995.32753952077
        },
        {
          "locationType": "Commercial Hub",
          "density": "M",
          "division": "Khulna",
          "key": "Commercial HubMKhulna",
          "monthlySales": 1359499.2,
          "monthlyFootfall": 3361.2,
          "basket": 404.4684041413781,
          "profit": 58187.884841119245
        },
        {
          "locationType": "Commercial Hub",
          "density": "M",
          "division": "Mymensingh",
          "key": "Commercial HubMMymensingh",
          "monthlySales": 1469241.5,
          "monthlyFootfall": 3572.5,
          "basket": 411.26424072778167,
          "profit": 55847.57154692778
        },
        {
          "locationType": "Commercial Hub",
          "density": "M",
          "division": "Rajshahi",
          "key": "Commercial HubMRajshahi",
          "monthlySales": 1462370,
          "monthlyFootfall": 3678.4,
          "basket": 397.55600260983033,
          "profit": 66256.31919749579
        },
        {
          "locationType": "Commercial Hub",
          "density": "M",
          "division": "Rangpur",
          "key": "Commercial HubMRangpur",
          "monthlySales": 1477614.5882352942,
          "monthlyFootfall": 3406.4117647058824,
          "basket": 433.7745082802328,
          "profit": 67468.7817661952
        },
        {
          "locationType": "Commercial Hub",
          "density": "M",
          "division": "Sylhet",
          "key": "Commercial HubMSylhet",
          "monthlySales": 1219248.6666666667,
          "monthlyFootfall": 3243.1111111111113,
          "basket": 375.9503220501576,
          "profit": 14704.106321562715
        },
        {
          "locationType": "Gated Community",
          "density": "H",
          "division": "Dhaka",
          "key": "Gated CommunityHDhaka",
          "monthlySales": 8554645.5,
          "monthlyFootfall": 13233.5,
          "basket": 646.4386216798277,
          "profit": -75236.9918476291
        },
        {
          "locationType": "Gated Community",
          "density": "H",
          "division": "DhakaGBUD",
          "key": "Gated CommunityHDhakaGBUD",
          "monthlySales": 40720005,
          "monthlyFootfall": 33692,
          "basket": 1208.5956606909651,
          "profit": 2746969.340041858
        },
        {
          "locationType": "Gated Community",
          "density": "L",
          "division": "Dhaka",
          "key": "Gated CommunityLDhaka",
          "monthlySales": 2068044,
          "monthlyFootfall": 4460,
          "basket": 463.68699551569506,
          "profit": 68237.75110719487
        },
        {
          "locationType": "Gated Community",
          "density": "L",
          "division": "Khulna",
          "key": "Gated CommunityLKhulna",
          "monthlySales": 884554,
          "monthlyFootfall": 3362,
          "basket": 263.10350981558594,
          "profit": 67863.5669112033
        },
        {
          "locationType": "Gated Community",
          "density": "M",
          "division": "Dhaka",
          "key": "Gated CommunityMDhaka",
          "monthlySales": 19828573.5,
          "monthlyFootfall": 24044,
          "basket": 824.6786516386625,
          "profit": 584939.2158392939
        },
        {
          "locationType": "High Street",
          "density": "H",
          "division": "Dhaka",
          "key": "High StreetHDhaka",
          "monthlySales": 1762574,
          "monthlyFootfall": 4361,
          "basket": 404.1673927998166,
          "profit": -17408.75687877403
        },
        {
          "locationType": "High Street",
          "density": "L",
          "division": "Chattogram",
          "key": "High StreetLChattogram",
          "monthlySales": 927574.5,
          "monthlyFootfall": 2164.5,
          "basket": 428.53984753984753,
          "profit": -55579.577316935975
        },
        {
          "locationType": "High Street",
          "density": "L",
          "division": "Dhaka",
          "key": "High StreetLDhaka",
          "monthlySales": 998709.1111111111,
          "monthlyFootfall": 2579.3333333333335,
          "basket": 387.196605496683,
          "profit": 24788.493660413118
        },
        {
          "locationType": "High Street",
          "density": "L",
          "division": "Khulna",
          "key": "High StreetLKhulna",
          "monthlySales": 580677,
          "monthlyFootfall": 1573,
          "basket": 369.15257469802924,
          "profit": -29824.277681775508
        },
        {
          "locationType": "High Street",
          "density": "L",
          "division": "Mymensingh",
          "key": "High StreetLMymensingh",
          "monthlySales": 470393,
          "monthlyFootfall": 1233,
          "basket": 381.5028386050284,
          "profit": -30047.053828680422
        },
        {
          "locationType": "High Street",
          "density": "L",
          "division": "Rangpur",
          "key": "High StreetLRangpur",
          "monthlySales": 696766.5,
          "monthlyFootfall": 1820,
          "basket": 382.8387362637363,
          "profit": -46650.24337194754
        },
        {
          "locationType": "High Street",
          "density": "M",
          "division": "Barishal",
          "key": "High StreetMBarishal",
          "monthlySales": 790941.25,
          "monthlyFootfall": 2419.5,
          "basket": 326.9027691671833,
          "profit": -38125.46216276294
        },
        {
          "locationType": "High Street",
          "density": "M",
          "division": "Chattogram",
          "key": "High StreetMChattogram",
          "monthlySales": 856523,
          "monthlyFootfall": 1817,
          "basket": 471.3940561364887,
          "profit": 17274.248631063674
        },
        {
          "locationType": "High Street",
          "density": "M",
          "division": "Dhaka",
          "key": "High StreetMDhaka",
          "monthlySales": 1309425.4545454546,
          "monthlyFootfall": 3146.5454545454545,
          "basket": 416.1470010401017,
          "profit": 52266.98968266524
        },
        {
          "locationType": "High Street",
          "density": "M",
          "division": "Khulna",
          "key": "High StreetMKhulna",
          "monthlySales": 1062973.5714285714,
          "monthlyFootfall": 2635.5714285714284,
          "basket": 403.31806601983845,
          "profit": -4300.297415296115
        },
        {
          "locationType": "High Street",
          "density": "M",
          "division": "Mymensingh",
          "key": "High StreetMMymensingh",
          "monthlySales": 829897,
          "monthlyFootfall": 1944.5,
          "basket": 426.7919773720751,
          "profit": -9871.926777520792
        },
        {
          "locationType": "High Street",
          "density": "M",
          "division": "Rajshahi",
          "key": "High StreetMRajshahi",
          "monthlySales": 894370,
          "monthlyFootfall": 2385,
          "basket": 374.9979035639413,
          "profit": 13598.17549183956
        },
        {
          "locationType": "High Street",
          "density": "M",
          "division": "Rangpur",
          "key": "High StreetMRangpur",
          "monthlySales": 880542.8,
          "monthlyFootfall": 2418,
          "basket": 364.16162117452444,
          "profit": 19487.965603912242
        },
        {
          "locationType": "High Street",
          "density": "M",
          "division": "Sylhet",
          "key": "High StreetMSylhet",
          "monthlySales": 1203666.5,
          "monthlyFootfall": 2955.5,
          "basket": 407.26323803079003,
          "profit": 31423.2798639292
        },
        {
          "locationType": "High Street W Institutions",
          "density": "H",
          "division": "Barishal",
          "key": "High Street W InstitutionsHBarishal",
          "monthlySales": 2421928,
          "monthlyFootfall": 6737,
          "basket": 359.49651180050466,
          "profit": 116204.36814952362
        },
        {
          "locationType": "High Street W Institutions",
          "density": "H",
          "division": "Dhaka",
          "key": "High Street W InstitutionsHDhaka",
          "monthlySales": 7279102,
          "monthlyFootfall": 12882,
          "basket": 565.0599285825183,
          "profit": 777179.2645156041
        },
        {
          "locationType": "High Street W Institutions",
          "density": "H",
          "division": "Khulna",
          "key": "High Street W InstitutionsHKhulna",
          "monthlySales": 5359746,
          "monthlyFootfall": 9805,
          "basket": 546.6339622641509,
          "profit": 174270.54410812235
        },
        {
          "locationType": "High Street W Institutions",
          "density": "L",
          "division": "Barishal",
          "key": "High Street W InstitutionsLBarishal",
          "monthlySales": 1206186,
          "monthlyFootfall": 3243.75,
          "basket": 371.8492485549133,
          "profit": 50169.69288189825
        },
        {
          "locationType": "High Street W Institutions",
          "density": "L",
          "division": "Chattogram",
          "key": "High Street W InstitutionsLChattogram",
          "monthlySales": 1310365.5,
          "monthlyFootfall": 2932,
          "basket": 446.918656207367,
          "profit": 52037.083173775536
        },
        {
          "locationType": "High Street W Institutions",
          "density": "L",
          "division": "Dhaka",
          "key": "High Street W InstitutionsLDhaka",
          "monthlySales": 1344656.6666666667,
          "monthlyFootfall": 3575.3333333333335,
          "basket": 376.09267201193364,
          "profit": 70871.09946117159
        },
        {
          "locationType": "High Street W Institutions",
          "density": "L",
          "division": "Mymensingh",
          "key": "High Street W InstitutionsLMymensingh",
          "monthlySales": 1650204,
          "monthlyFootfall": 2937,
          "basket": 561.8672114402451,
          "profit": 120778.6028013252
        },
        {
          "locationType": "High Street W Institutions",
          "density": "L",
          "division": "Rajshahi",
          "key": "High Street W InstitutionsLRajshahi",
          "monthlySales": 1194429,
          "monthlyFootfall": 3199.5,
          "basket": 373.3173933427098,
          "profit": 49452.865204380316
        },
        {
          "locationType": "High Street W Institutions",
          "density": "M",
          "division": "Barishal",
          "key": "High Street W InstitutionsMBarishal",
          "monthlySales": 1399515,
          "monthlyFootfall": 3037,
          "basket": 460.8215344089562,
          "profit": 31308.014868472896
        },
        {
          "locationType": "High Street W Institutions",
          "density": "M",
          "division": "Dhaka",
          "key": "High Street W InstitutionsMDhaka",
          "monthlySales": 2853911.8,
          "monthlyFootfall": 6337.2,
          "basket": 450.3427065581013,
          "profit": 171521.63348264544
        },
        {
          "locationType": "High Street W Institutions",
          "density": "M",
          "division": "Khulna",
          "key": "High Street W InstitutionsMKhulna",
          "monthlySales": 2385030,
          "monthlyFootfall": 5403,
          "basket": 441.4269850083287,
          "profit": 135868.89904502052
        },
        {
          "locationType": "High Street W Institutions",
          "density": "M",
          "division": "Rajshahi",
          "key": "High Street W InstitutionsMRajshahi",
          "monthlySales": 1713017,
          "monthlyFootfall": 4382,
          "basket": 390.9212688270196,
          "profit": 87031.45754498593
        },
        {
          "locationType": "High Street W Institutions",
          "density": "M",
          "division": "Rangpur",
          "key": "High Street W InstitutionsMRangpur",
          "monthlySales": 1438581,
          "monthlyFootfall": 3694.5,
          "basket": 389.3844904587901,
          "profit": 47919.631605405404
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "H",
          "division": "Barishal",
          "key": "High Street W Residential BlockHBarishal",
          "monthlySales": 2382223.8,
          "monthlyFootfall": 5088,
          "basket": 468.20436320754715,
          "profit": 108254.10462767424
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "H",
          "division": "Chattogram",
          "key": "High Street W Residential BlockHChattogram",
          "monthlySales": 7079980.176470588,
          "monthlyFootfall": 9446.058823529413,
          "basket": 749.5168417578448,
          "profit": 628322.4179279655
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "H",
          "division": "Dhaka",
          "key": "High Street W Residential BlockHDhaka",
          "monthlySales": 7383048.360655738,
          "monthlyFootfall": 10379.622950819672,
          "basket": 711.3021730787151,
          "profit": 590132.0124207502
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "H",
          "division": "DhakaGBUD",
          "key": "High Street W Residential BlockHDhakaGBUD",
          "monthlySales": 25843934.2,
          "monthlyFootfall": 23948.666666666668,
          "basket": 1079.1387495476436,
          "profit": 2464503.67453852
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "H",
          "division": "Khulna",
          "key": "High Street W Residential BlockHKhulna",
          "monthlySales": 4246169.857142857,
          "monthlyFootfall": 7275.571428571428,
          "basket": 583.6201182037739,
          "profit": 198316.09290350747
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "H",
          "division": "Mymensingh",
          "key": "High Street W Residential BlockHMymensingh",
          "monthlySales": 3767816,
          "monthlyFootfall": 7260,
          "basket": 518.9829201101928,
          "profit": 236953.3583198465
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "H",
          "division": "Rajshahi",
          "key": "High Street W Residential BlockHRajshahi",
          "monthlySales": 3131916.75,
          "monthlyFootfall": 4982,
          "basket": 628.646477318346,
          "profit": 164278.52319697101
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "H",
          "division": "Rangpur",
          "key": "High Street W Residential BlockHRangpur",
          "monthlySales": 11189254.5,
          "monthlyFootfall": 13464,
          "basket": 831.0497994652407,
          "profit": 689628.8321860823
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "H",
          "division": "Sylhet",
          "key": "High Street W Residential BlockHSylhet",
          "monthlySales": 7088700.1,
          "monthlyFootfall": 9237.4,
          "basket": 767.3912681057441,
          "profit": 551687.6792437853
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "L",
          "division": "Barishal",
          "key": "High Street W Residential BlockLBarishal",
          "monthlySales": 947764.0666666667,
          "monthlyFootfall": 2615.133333333333,
          "basket": 362.4151987151707,
          "profit": 15248.238204849795
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "L",
          "division": "Chattogram",
          "key": "High Street W Residential BlockLChattogram",
          "monthlySales": 1210110.7692307692,
          "monthlyFootfall": 2710.3076923076924,
          "basket": 446.4846455128569,
          "profit": 59804.66511954694
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "L",
          "division": "Dhaka",
          "key": "High Street W Residential BlockLDhaka",
          "monthlySales": 1344285.0476190476,
          "monthlyFootfall": 2921,
          "basket": 460.21398412155,
          "profit": 50478.8897929744
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "L",
          "division": "Khulna",
          "key": "High Street W Residential BlockLKhulna",
          "monthlySales": 1063608.3333333333,
          "monthlyFootfall": 2620,
          "basket": 405.95737913486005,
          "profit": 50477.637166559805
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "L",
          "division": "Mymensingh",
          "key": "High Street W Residential BlockLMymensingh",
          "monthlySales": 822515,
          "monthlyFootfall": 2231,
          "basket": 368.6754818467055,
          "profit": -15505.131961754552
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "L",
          "division": "Rajshahi",
          "key": "High Street W Residential BlockLRajshahi",
          "monthlySales": 1443903,
          "monthlyFootfall": 3363.285714285714,
          "basket": 429.31321411884636,
          "profit": 54825.69652081281
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "L",
          "division": "Rangpur",
          "key": "High Street W Residential BlockLRangpur",
          "monthlySales": 1131227.0833333333,
          "monthlyFootfall": 2787.0833333333335,
          "basket": 405.8820451487516,
          "profit": 37321.70835215843
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "L",
          "division": "Sylhet",
          "key": "High Street W Residential BlockLSylhet",
          "monthlySales": 1519857.25,
          "monthlyFootfall": 3665,
          "basket": 414.6950204638472,
          "profit": 67514.26886278336
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "M",
          "division": "Barishal",
          "key": "High Street W Residential BlockMBarishal",
          "monthlySales": 1194110.05,
          "monthlyFootfall": 3186.95,
          "basket": 374.68741273003974,
          "profit": 29679.32779404478
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "M",
          "division": "Chattogram",
          "key": "High Street W Residential BlockMChattogram",
          "monthlySales": 1936744.28125,
          "monthlyFootfall": 4025.59375,
          "basket": 481.10773255498026,
          "profit": 115046.68006186208
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "M",
          "division": "Dhaka",
          "key": "High Street W Residential BlockMDhaka",
          "monthlySales": 2877459.51754386,
          "monthlyFootfall": 5403.833333333333,
          "basket": 532.4848750967881,
          "profit": 146681.16030403992
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "M",
          "division": "DhakaGBUD",
          "key": "High Street W Residential BlockMDhakaGBUD",
          "monthlySales": 1561111,
          "monthlyFootfall": 3806,
          "basket": 410.1710457172885,
          "profit": 32541.22912455842
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "M",
          "division": "Khulna",
          "key": "High Street W Residential BlockMKhulna",
          "monthlySales": 1765883.5238095238,
          "monthlyFootfall": 3838.5714285714284,
          "basket": 460.0366455774718,
          "profit": 94311.8888477782
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "M",
          "division": "Mymensingh",
          "key": "High Street W Residential BlockMMymensingh",
          "monthlySales": 1404443.5,
          "monthlyFootfall": 3435.5,
          "basket": 408.8032309707466,
          "profit": 53653.08609692082
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "M",
          "division": "Rajshahi",
          "key": "High Street W Residential BlockMRajshahi",
          "monthlySales": 1789165.1818181819,
          "monthlyFootfall": 3817,
          "basket": 468.7359658942054,
          "profit": 122451.61565035797
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "M",
          "division": "Rangpur",
          "key": "High Street W Residential BlockMRangpur",
          "monthlySales": 1422138.4210526317,
          "monthlyFootfall": 3546.5263157894738,
          "basket": 400.9947465273656,
          "profit": 50376.92321644047
        },
        {
          "locationType": "High Street W Residential Block",
          "density": "M",
          "division": "Sylhet",
          "key": "High Street W Residential BlockMSylhet",
          "monthlySales": 2010484.4583333333,
          "monthlyFootfall": 4117.75,
          "basket": 488.2483051018962,
          "profit": 103505.25417353335
        },
        {
          "locationType": "Industrial",
          "density": "H",
          "division": "Chattogram",
          "key": "IndustrialHChattogram",
          "monthlySales": 11313587,
          "monthlyFootfall": 14444,
          "basket": 783.2724314594295,
          "profit": 1149666.5170246214
        },
        {
          "locationType": "Industrial",
          "density": "H",
          "division": "Dhaka",
          "key": "IndustrialHDhaka",
          "monthlySales": 3541681,
          "monthlyFootfall": 7955,
          "basket": 445.2144563167819,
          "profit": 308604.5000815202
        },
        {
          "locationType": "Industrial",
          "density": "H",
          "division": "Mymensingh",
          "key": "IndustrialHMymensingh",
          "monthlySales": 3048746,
          "monthlyFootfall": 6451,
          "basket": 472.60052705006973,
          "profit": 111608.63444655726
        },
        {
          "locationType": "Industrial",
          "density": "L",
          "division": "Dhaka",
          "key": "IndustrialLDhaka",
          "monthlySales": 2000287,
          "monthlyFootfall": 3882.3333333333335,
          "basket": 515.2280415557655,
          "profit": 119734.92178106343
        },
        {
          "locationType": "Industrial",
          "density": "L",
          "division": "Sylhet",
          "key": "IndustrialLSylhet",
          "monthlySales": 1703067,
          "monthlyFootfall": 3124,
          "basket": 545.1558898847632,
          "profit": 92350.14595061989
        },
        {
          "locationType": "Industrial",
          "density": "M",
          "division": "Dhaka",
          "key": "IndustrialMDhaka",
          "monthlySales": 2028151.6666666667,
          "monthlyFootfall": 4416.333333333333,
          "basket": 459.2388104762624,
          "profit": 101908.01559273146
        },
        {
          "locationType": "Inside the Mall",
          "density": "H",
          "division": "Dhaka",
          "key": "Inside the MallHDhaka",
          "monthlySales": 7599044,
          "monthlyFootfall": 11301,
          "basket": 672.422263516503,
          "profit": 897520.7511309914
        },
        {
          "locationType": "Inside the Mall",
          "density": "H",
          "division": "Mymensingh",
          "key": "Inside the MallHMymensingh",
          "monthlySales": 5427417,
          "monthlyFootfall": 10464,
          "basket": 518.6751720183486,
          "profit": 356245.0960102
        },
        {
          "locationType": "Inside the Mall",
          "density": "H",
          "division": "Rangpur",
          "key": "Inside the MallHRangpur",
          "monthlySales": 4688844,
          "monthlyFootfall": 7072,
          "basket": 663.0152714932127,
          "profit": 282102.6901911871
        },
        {
          "locationType": "Inside the Mall",
          "density": "L",
          "division": "Khulna",
          "key": "Inside the MallLKhulna",
          "monthlySales": 1342148,
          "monthlyFootfall": 3668,
          "basket": 365.907306434024,
          "profit": 90679.30813128862
        },
        {
          "locationType": "Inside the Mall",
          "density": "L",
          "division": "Rajshahi",
          "key": "Inside the MallLRajshahi",
          "monthlySales": 1160626,
          "monthlyFootfall": 1971,
          "basket": 588.8513444951801,
          "profit": 19844.928112671827
        },
        {
          "locationType": "Inside the Mall",
          "density": "M",
          "division": "Barishal",
          "key": "Inside the MallMBarishal",
          "monthlySales": 1077061,
          "monthlyFootfall": 2785.5,
          "basket": 386.6670256686412,
          "profit": 32760.71595146606
        },
        {
          "locationType": "Inside the Mall",
          "density": "M",
          "division": "Dhaka",
          "key": "Inside the MallMDhaka",
          "monthlySales": 2063756,
          "monthlyFootfall": 4714,
          "basket": 437.7929571489181,
          "profit": 113077.53596727249
        },
        {
          "locationType": "Inside the Mall",
          "density": "M",
          "division": "Rangpur",
          "key": "Inside the MallMRangpur",
          "monthlySales": 2495247,
          "monthlyFootfall": 5037.5,
          "basket": 495.33439205955335,
          "profit": 204157.8516485553
        },
        {
          "locationType": "Within the Mahalla",
          "density": "H",
          "division": "Chattogram",
          "key": "Within the MahallaHChattogram",
          "monthlySales": 1001878.6666666666,
          "monthlyFootfall": 2494.6666666666665,
          "basket": 401.6082308925708,
          "profit": 10729.38640561481
        },
        {
          "locationType": "Within the Mahalla",
          "density": "H",
          "division": "Dhaka",
          "key": "Within the MahallaHDhaka",
          "monthlySales": 2850894.447368421,
          "monthlyFootfall": 5549.4473684210525,
          "basket": 513.725828555712,
          "profit": 132393.00803261442
        },
        {
          "locationType": "Within the Mahalla",
          "density": "H",
          "division": "DhakaGBUD",
          "key": "Within the MahallaHDhakaGBUD",
          "monthlySales": 1608280,
          "monthlyFootfall": 4260,
          "basket": 377.53051643192487,
          "profit": -66070.97816834765
        },
        {
          "locationType": "Within the Mahalla",
          "density": "H",
          "division": "Mymensingh",
          "key": "Within the MahallaHMymensingh",
          "monthlySales": 1497162,
          "monthlyFootfall": 5065,
          "basket": 295.5897334649556,
          "profit": -6657.479511355154
        },
        {
          "locationType": "Within the Mahalla",
          "density": "H",
          "division": "Rajshahi",
          "key": "Within the MahallaHRajshahi",
          "monthlySales": 372966,
          "monthlyFootfall": 929,
          "basket": 401.470398277718,
          "profit": -50438.53217862311
        },
        {
          "locationType": "Within the Mahalla",
          "density": "L",
          "division": "Barishal",
          "key": "Within the MahallaLBarishal",
          "monthlySales": 690815,
          "monthlyFootfall": 2059,
          "basket": 335.5099562894609,
          "profit": 6131.6785155877005
        },
        {
          "locationType": "Within the Mahalla",
          "density": "L",
          "division": "Chattogram",
          "key": "Within the MahallaLChattogram",
          "monthlySales": 1157919,
          "monthlyFootfall": 2056,
          "basket": 563.1901750972762,
          "profit": 24908.75532669676
        },
        {
          "locationType": "Within the Mahalla",
          "density": "L",
          "division": "Dhaka",
          "key": "Within the MahallaLDhaka",
          "monthlySales": 1323465.3,
          "monthlyFootfall": 2919.95,
          "basket": 453.2493022140791,
          "profit": 48806.275845114826
        },
        {
          "locationType": "Within the Mahalla",
          "density": "L",
          "division": "Khulna",
          "key": "Within the MahallaLKhulna",
          "monthlySales": 867594,
          "monthlyFootfall": 2108,
          "basket": 411.5721062618596,
          "profit": 28006.11376124411
        },
        {
          "locationType": "Within the Mahalla",
          "density": "L",
          "division": "Mymensingh",
          "key": "Within the MahallaLMymensingh",
          "monthlySales": 712066,
          "monthlyFootfall": 1917,
          "basket": 371.44809598330727,
          "profit": -5210.810687398276
        },
        {
          "locationType": "Within the Mahalla",
          "density": "L",
          "division": "Rajshahi",
          "key": "Within the MahallaLRajshahi",
          "monthlySales": 955707.5,
          "monthlyFootfall": 2604.5,
          "basket": 366.9447110769822,
          "profit": 27036.198983782815
        },
        {
          "locationType": "Within the Mahalla",
          "density": "L",
          "division": "Rangpur",
          "key": "Within the MahallaLRangpur",
          "monthlySales": 943068,
          "monthlyFootfall": 2212,
          "basket": 426.34177215189874,
          "profit": 14906.533307265112
        },
        {
          "locationType": "Within the Mahalla",
          "density": "M",
          "division": "Barishal",
          "key": "Within the MahallaMBarishal",
          "monthlySales": 1045632,
          "monthlyFootfall": 3425,
          "basket": 305.2940145985402,
          "profit": 10549.135227464052
        },
        {
          "locationType": "Within the Mahalla",
          "density": "M",
          "division": "Chattogram",
          "key": "Within the MahallaMChattogram",
          "monthlySales": 1234851.7142857143,
          "monthlyFootfall": 3302.714285714286,
          "basket": 373.8899606384359,
          "profit": 56493.133042871945
        },
        {
          "locationType": "Within the Mahalla",
          "density": "M",
          "division": "Dhaka",
          "key": "Within the MahallaMDhaka",
          "monthlySales": 1852949.7738095238,
          "monthlyFootfall": 4244.166666666667,
          "basket": 436.58741970772206,
          "profit": 63487.49470141468
        },
        {
          "locationType": "Within the Mahalla",
          "density": "M",
          "division": "Khulna",
          "key": "Within the MahallaMKhulna",
          "monthlySales": 1442189.2,
          "monthlyFootfall": 3266,
          "basket": 441.5766074709124,
          "profit": 56727.873773508414
        },
        {
          "locationType": "Within the Mahalla",
          "density": "M",
          "division": "Rajshahi",
          "key": "Within the MahallaMRajshahi",
          "monthlySales": 1297604,
          "monthlyFootfall": 3267,
          "basket": 397.18518518518516,
          "profit": 30326.73980873922
        },
        {
          "locationType": "Within the Mahalla",
          "density": "M",
          "division": "Rangpur",
          "key": "Within the MahallaMRangpur",
          "monthlySales": 1460791.5,
          "monthlyFootfall": 3528.3333333333335,
          "basket": 414.0174303259329,
          "profit": 22961.876987966232
        }
      ]
    },
    "gpLookup": {
      "Chattogram FR Non-PNP": 0.1652883053847348,
      "Dhaka FR Non-PNP": 0.15969475504252637,
      "Barishal FR Non-PNP": 0.1560199267819153,
      "Mymensingh FR Non-PNP": 0.15572779254330107,
      "Sylhet FR PNP": 0.14642072725733013,
      "Chattogram FR PNP": 0.17487179995671584,
      "Rajshahi FR Non-PNP": 0.15744785972112785,
      "Khulna FR Non-PNP": 0.16189872955754886,
      "Rangpur FR PNP": 0.1591380946046694,
      "Rangpur FR Non-PNP": 0.15282818453659383,
      "Khulna FR PNP": 0.14902038255368039,
      "Rajshahi OWN Non-PNP": 0.1779735306579332,
      "Dhaka FR PNP": 0.1576680362008356,
      "Rajshahi FR PNP": 0.16,
      "Sylhet FR Non-PNP": 0.15249914700152678,
      "Barishal OWN PNP": 0.15686241014117036,
      "Chattogram OWN PNP": 0.16384349743590615,
      "Chattogram OWN Non-PNP": 0.1872357859116617,
      "Dhaka OWN PNP": 0.1675076152205274,
      "Dhaka GBUD OWN PNP": 0.17314744692529424,
      "Dhaka OWN Non-PNP": 0.16518961505295704,
      "Khulna OWN Non-PNP": 0.15726401257003114,
      "Khulna OWN PNP": 0.14245785152560209,
      "Rajshahi OWN PNP": 0.15853756850791623,
      "Sylhet OWN PNP": 0.1580248452828885
    },
    "basketLookup": {
      "Chattogram FR Non-PNP": 485.0897135593342,
      "Dhaka FR Non-PNP": 426.75836724362256,
      "Barishal FR Non-PNP": 416.7793579929576,
      "Mymensingh FR Non-PNP": 418.871508662188,
      "Sylhet FR PNP": 451.9897896311944,
      "Chattogram FR PNP": 573.290215899164,
      "Rajshahi FR Non-PNP": 440.139844582315,
      "Khulna FR Non-PNP": 410.652048994405,
      "Rangpur FR PNP": 820.181925115314,
      "Rangpur FR Non-PNP": 427.3099146360964,
      "Khulna FR PNP": 726.3033339948403,
      "Rajshahi OWN Non-PNP": 531.428588915002,
      "Dhaka FR PNP": 414.02385313099603,
      "Rajshahi FR PNP": 368.0476673427992,
      "Sylhet FR Non-PNP": 277.6853658536586,
      "Barishal OWN PNP": 565.5688301228271,
      "Chattogram OWN PNP": 865.0753001159181,
      "Chattogram OWN Non-PNP": 492.6432162875042,
      "Dhaka OWN PNP": 722.7058610576797,
      "Dhaka GBUD OWN PNP": 1006.6521467414644,
      "Dhaka OWN Non-PNP": 487.8810581890003,
      "Khulna OWN Non-PNP": 602.3142129304288,
      "Khulna OWN PNP": 524.9264072990188,
      "Rajshahi OWN PNP": 635.3894739054585,
      "Sylhet OWN PNP": 774.6384718215462
    },
    "writeOffLookup": {
      "Chattogram FR Non-PNP": 0.004696693094894034,
      "Dhaka FR Non-PNP": 0.004840764918852674,
      "Barishal FR Non-PNP": 0.004123813019941308,
      "Mymensingh FR Non-PNP": 0.004195758518699045,
      "Sylhet FR PNP": 0.004443518483970169,
      "Chattogram FR PNP": 0.0033987884994865127,
      "Rajshahi FR Non-PNP": 0.0065464263782539734,
      "Khulna FR Non-PNP": 0.005060823427871265,
      "Rangpur FR PNP": 0.00398549978019622,
      "Rangpur FR Non-PNP": 0.004951331799037794,
      "Khulna FR PNP": 0.004041028026122809,
      "Dhaka FR PNP": 0.007449372170410415,
      "Rajshahi FR PNP": 0.0055,
      "Sylhet FR Non-PNP": 0.00498941390671811,
      "Barishal OWN PNP": 0.006193408249442319,
      "Chattogram OWN PNP": 0.004434376955230566,
      "Chattogram OWN Non-PNP": 0.004643033707457255,
      "Dhaka OWN PNP": 0.004757485979176226,
      "Dhaka GBUD OWN PNP": 0.005,
      "Dhaka OWN Non-PNP": 0.0038256335464001753,
      "Khulna OWN Non-PNP": 0.003492621992306228,
      "Khulna OWN PNP": 0.0035921925142926316,
      "Rajshahi OWN PNP": 0.005524235952109365,
      "Sylhet OWN PNP": 0.005066693493728824
    },
    "autoGpPercent": 0.15969475504252637,
    "autoGpShareFr": 0.4,
    "franchiseCommissionTable": [
      [
        0.4,
        216236094.96396768,
        81823342.31780863,
        0.3783981685918033
      ],
      [
        0.5,
        14612129.16419092,
        6923298.733507378,
        0.4738049230001263
      ],
      [
        0.45,
        19106519.396355767,
        8179967.367512893,
        0.42812441124536155
      ],
      [
        0.35,
        2944504.8046399434,
        979987.2382083233,
        0.33281903179918804
      ],
      [
        0.43,
        1976743.1024383586,
        800461.8044278886,
        0.4049397230426657
      ],
      [
        0.36,
        198867.13133755504,
        67806.72577428953,
        0.34096497152762306
      ],
      [
        0.38,
        1148509.1020809629,
        409356.10384340404,
        0.3564239091372451
      ],
      [
        0.37,
        620666.9914418305,
        218792.5786446599,
        0.35251202603250625
      ],
      [
        0.42,
        882030.8621041488,
        358912.5539556261,
        0.4069160948625018
      ]
    ],
    "referenceSalesPerDay": 119126.78125,
    "referenceFootfall": 217.24375,
    "referenceBasket": 548.3553899709427,
    "referenceProfit": 257639.5005627723,
    "autoBasketSize": 426.75836724362256,
    "categories": [
      {
        "name": "Baby Care",
        "mix": 0.013472604431973215
      },
      {
        "name": "Baby Food",
        "mix": 0.011715212035425736
      },
      {
        "name": "Beverage & Tobacco",
        "mix": 0.05279877768297659
      },
      {
        "name": "Dairy",
        "mix": 0.07117234807056318
      },
      {
        "name": "Home Care",
        "mix": 0.08023840843184718
      },
      {
        "name": "Kitchen Additives",
        "mix": 0.052601455715612
      },
      {
        "name": "Packaged Foods",
        "mix": 0.25734527058542767
      },
      {
        "name": "Personal Care",
        "mix": 0.1390845214702441
      },
      {
        "name": "Commodities",
        "mix": 0.27201846436334753
      },
      {
        "name": "Electronics & Appl",
        "mix": 0.010335709814317895
      },
      {
        "name": "Gift & Toys",
        "mix": 0.008444104467581955
      },
      {
        "name": "Home Accessories",
        "mix": 0.017688388729473422
      },
      {
        "name": "Home Appliance",
        "mix": 0.006835174345715071
      },
      {
        "name": "Stationeries",
        "mix": 0.0038038159329865026
      },
      {
        "name": "Life Style",
        "mix": 0.0011784602523802336
      },
      {
        "name": "Perishables",
        "mix": 0
      },
      {
        "name": "Protein",
        "mix": 0.001267283670127819
      }
    ]
  },
  "signatories": [
    {
      "role": "Prepared by",
      "name": "Md Aftab Ul Islam",
      "designation": "Senior Executive, Retail Operations",
      "signatureId": "source-signature-2",
      "includeInPdf": true
    },
    {
      "role": "Sales Given By",
      "name": "Mr. Mamun",
      "designation": "Leader's Name",
      "autoSource": "sales-given-by",
      "manualOverride": false,
      "signatureId": "source-signature-1",
      "includeInPdf": true
    },
    {
      "role": "Opened by",
      "name": "Mr. Abu Sayem Mohammad Al Azad",
      "designation": "Head Of Business Expansion",
      "autoSource": "opened-by",
      "manualOverride": false,
      "signatureId": "",
      "includeInPdf": false
    },
    {
      "role": "Viewed by",
      "name": "Salah Uddin Misbah",
      "designation": "Head Of Business, CG",
      "signatureId": "salah-uddin-misbah-signature",
      "includeInPdf": false
    },
    {
      "role": "Viewed by",
      "name": "Saiful Alam Rasel",
      "designation": "Assistant Director, Operations",
      "signatureId": "saiful-alam-rasel-signature",
      "includeInPdf": true
    },
    {
      "role": "Viewed by",
      "name": "Tamal Paul",
      "designation": "Director – Business Planning & Investor Relations",
      "signatureId": "tamal-paul-signature",
      "includeInPdf": false
    },
    {
      "role": "Approved By",
      "name": "Abu Naser",
      "designation": "Chief Business Officer",
      "signatureId": "",
      "includeInPdf": false
    }
  ]
};
