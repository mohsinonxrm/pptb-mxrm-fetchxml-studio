/**
 * FetchXmlToQueryExpression responses (the Query property), captured from a sandbox environment on
 * 2026-10-31 for the C# QueryExpression generator tests (#178). Read-only calls; no record data.
 * The FetchXML for each is in queryExpressionCodegen.test.ts.
 */
import type { QeQueryExpression } from "../queryExpressionTypes";

export const SANDBOX_QUERY_EXPRESSIONS = {
	"linkAndOrder": {
		"Distinct": true,
		"NoLock": false,
		"QueryHints": "",
		"ForceSeek": null,
		"DataSource": null,
		"EntityName": "account",
		"TopCount": 5,
		"SubQueryExpression": null,
		"PageInfo": {
			"PageNumber": 0,
			"Count": 0,
			"ReturnTotalRecordCount": false,
			"ReturnRecordCountLessThanMarv": false,
			"PagingCookie": null
		},
		"LinkEntities": [
			{
				"LinkFromAttributeName": "accountid",
				"LinkFromEntityName": "account",
				"LinkToEntityName": "contact",
				"LinkToAttributeName": "parentcustomerid",
				"JoinOperator": "LeftOuter",
				"EntityAlias": "c",
				"ForceSeek": null,
				"LinkCriteria": {
					"FilterOperator": "Or",
					"FilterHint": null,
					"IsQuickFindFilter": false,
					"AnyAllFilterLinkEntity": null,
					"Conditions": [
						{
							"EntityName": null,
							"CompareColumns": false,
							"AttributeName": "firstname",
							"Operator": "Equal",
							"Values": [
								{
									"Type": "System.String",
									"Value": "00123"
								}
							]
						},
						{
							"EntityName": null,
							"CompareColumns": false,
							"AttributeName": "lastname",
							"Operator": "Like",
							"Values": [
								{
									"Type": "System.String",
									"Value": "a\"b\\c%"
								}
							]
						}
					],
					"Filters": []
				},
				"LinkEntities": [],
				"Columns": {
					"AllColumns": false,
					"Columns": [
						"fullname"
					],
					"AttributeExpressions": []
				},
				"Orders": [
					{
						"AttributeName": "fullname",
						"OrderType": "Ascending",
						"Alias": null,
						"EntityName": null
					}
				]
			}
		],
		"Criteria": {
			"FilterOperator": "And",
			"FilterHint": null,
			"IsQuickFindFilter": false,
			"AnyAllFilterLinkEntity": null,
			"Conditions": [],
			"Filters": []
		},
		"Orders": [
			{
				"AttributeName": "name",
				"OrderType": "Descending",
				"Alias": null,
				"EntityName": null
			}
		],
		"ColumnSet": {
			"AllColumns": false,
			"Columns": [
				"name"
			],
			"AttributeExpressions": []
		}
	},
	"types": {
		"Distinct": false,
		"NoLock": false,
		"QueryHints": "",
		"ForceSeek": null,
		"DataSource": null,
		"EntityName": "account",
		"TopCount": null,
		"SubQueryExpression": null,
		"PageInfo": {
			"PageNumber": 1,
			"Count": 0,
			"ReturnTotalRecordCount": false,
			"ReturnRecordCountLessThanMarv": false,
			"PagingCookie": null
		},
		"LinkEntities": [],
		"Criteria": {
			"FilterOperator": "And",
			"FilterHint": null,
			"IsQuickFindFilter": false,
			"AnyAllFilterLinkEntity": null,
			"Conditions": [
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "numberofemployees",
					"Operator": "GreaterThan",
					"Values": [
						{
							"Type": "System.Int32",
							"Value": 10
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "revenue",
					"Operator": "GreaterEqual",
					"Values": [
						{
							"Type": "System.Decimal",
							"Value": 1000.5
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "donotemail",
					"Operator": "Equal",
					"Values": [
						{
							"Type": "System.Boolean",
							"Value": true
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "accountid",
					"Operator": "NotEqual",
					"Values": [
						{
							"Type": "System.Guid",
							"Value": "00000000-0000-0000-0000-000000000001"
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "createdon",
					"Operator": "OnOrAfter",
					"Values": [
						{
							"Type": "System.String",
							"Value": "2026-01-02T00:00:00+00:00"
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "modifiedon",
					"Operator": "LessThan",
					"Values": [
						{
							"Type": "System.String",
							"Value": "2026-01-02T10:00:00+00:00"
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "industrycode",
					"Operator": "In",
					"Values": [
						{
							"Type": "System.Int32",
							"Value": 1
						},
						{
							"Type": "System.Int32",
							"Value": 2
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "numberofemployees",
					"Operator": "Between",
					"Values": [
						{
							"Type": "System.Int32",
							"Value": 1
						},
						{
							"Type": "System.Int32",
							"Value": 9
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "telephone1",
					"Operator": "Null",
					"Values": []
				},
				{
					"EntityName": null,
					"CompareColumns": true,
					"AttributeName": "name",
					"Operator": "NotEqual",
					"Values": [
						{
							"Type": "System.String",
							"Value": "accountnumber"
						}
					]
				}
			],
			"Filters": []
		},
		"Orders": [],
		"ColumnSet": {
			"AllColumns": false,
			"Columns": [
				"name"
			],
			"AttributeExpressions": []
		}
	},
	"dateOps": {
		"Distinct": false,
		"NoLock": false,
		"QueryHints": "",
		"ForceSeek": null,
		"DataSource": null,
		"EntityName": "account",
		"TopCount": null,
		"SubQueryExpression": null,
		"PageInfo": {
			"PageNumber": 1,
			"Count": 0,
			"ReturnTotalRecordCount": false,
			"ReturnRecordCountLessThanMarv": false,
			"PagingCookie": null
		},
		"LinkEntities": [],
		"Criteria": {
			"FilterOperator": "And",
			"FilterHint": null,
			"IsQuickFindFilter": false,
			"AnyAllFilterLinkEntity": null,
			"Conditions": [
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "createdon",
					"Operator": "LastXDays",
					"Values": [
						{
							"Type": "System.Int32",
							"Value": 7
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "createdon",
					"Operator": "OlderThanXMonths",
					"Values": [
						{
							"Type": "System.Int32",
							"Value": 3
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "createdon",
					"Operator": "ThisYear",
					"Values": []
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "ownerid",
					"Operator": "EqualUserId",
					"Values": []
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "owningbusinessunit",
					"Operator": "EqualBusinessId",
					"Values": []
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "ownerid",
					"Operator": "EqualUserOrUserTeams",
					"Values": []
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "createdon",
					"Operator": "OlderThanXDays",
					"Values": [
						{
							"Type": "System.Int32",
							"Value": 2
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "name",
					"Operator": "BeginsWith",
					"Values": [
						{
							"Type": "System.String",
							"Value": "A"
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "name",
					"Operator": "DoesNotEndWith",
					"Values": [
						{
							"Type": "System.String",
							"Value": "z"
						}
					]
				}
			],
			"Filters": []
		},
		"Orders": [],
		"ColumnSet": {
			"AllColumns": false,
			"Columns": [
				"name"
			],
			"AttributeExpressions": []
		}
	},
	"anyAll": {
		"Distinct": false,
		"NoLock": false,
		"QueryHints": "",
		"ForceSeek": null,
		"DataSource": null,
		"EntityName": "contact",
		"TopCount": null,
		"SubQueryExpression": null,
		"PageInfo": {
			"PageNumber": 1,
			"Count": 0,
			"ReturnTotalRecordCount": false,
			"ReturnRecordCountLessThanMarv": false,
			"PagingCookie": null
		},
		"LinkEntities": [],
		"Criteria": {
			"FilterOperator": "And",
			"FilterHint": null,
			"IsQuickFindFilter": false,
			"AnyAllFilterLinkEntity": null,
			"Conditions": [],
			"Filters": [
				{
					"FilterOperator": "Or",
					"FilterHint": null,
					"IsQuickFindFilter": false,
					"AnyAllFilterLinkEntity": null,
					"Conditions": [
						{
							"EntityName": null,
							"CompareColumns": false,
							"AttributeName": "statecode",
							"Operator": "Equal",
							"Values": [
								{
									"Type": "System.Int32",
									"Value": 1
								}
							]
						}
					],
					"Filters": [
						{
							"FilterOperator": "Or",
							"FilterHint": null,
							"IsQuickFindFilter": false,
							"AnyAllFilterLinkEntity": {
								"LinkFromAttributeName": "contactid",
								"LinkFromEntityName": "contact",
								"LinkToEntityName": "account",
								"LinkToAttributeName": "primarycontactid",
								"JoinOperator": "Any",
								"EntityAlias": null,
								"ForceSeek": null,
								"LinkCriteria": {
									"FilterOperator": "And",
									"FilterHint": null,
									"IsQuickFindFilter": false,
									"AnyAllFilterLinkEntity": null,
									"Conditions": [
										{
											"EntityName": null,
											"CompareColumns": false,
											"AttributeName": "name",
											"Operator": "Equal",
											"Values": [
												{
													"Type": "System.String",
													"Value": "Contoso"
												}
											]
										}
									],
									"Filters": []
								},
								"LinkEntities": [],
								"Columns": {
									"AllColumns": false,
									"Columns": [],
									"AttributeExpressions": []
								},
								"Orders": []
							},
							"Conditions": [],
							"Filters": []
						}
					]
				}
			]
		},
		"Orders": [],
		"ColumnSet": {
			"AllColumns": false,
			"Columns": [
				"fullname"
			],
			"AttributeExpressions": []
		}
	},
	"aggregate": {
		"Distinct": false,
		"NoLock": false,
		"QueryHints": "",
		"ForceSeek": null,
		"DataSource": null,
		"EntityName": "account",
		"TopCount": null,
		"SubQueryExpression": null,
		"PageInfo": {
			"PageNumber": 1,
			"Count": 0,
			"ReturnTotalRecordCount": false,
			"ReturnRecordCountLessThanMarv": false,
			"PagingCookie": null
		},
		"LinkEntities": [],
		"Criteria": {
			"FilterOperator": "And",
			"FilterHint": null,
			"IsQuickFindFilter": false,
			"AnyAllFilterLinkEntity": null,
			"Conditions": [],
			"Filters": []
		},
		"Orders": [
			{
				"AttributeName": "address1_city",
				"OrderType": "Ascending",
				"Alias": "city",
				"EntityName": null
			}
		],
		"ColumnSet": {
			"AllColumns": false,
			"Columns": [],
			"AttributeExpressions": [
				{
					"AttributeName": "numberofemployees",
					"AggregateType": "Sum",
					"Alias": "total",
					"HasGroupBy": false,
					"DateTimeGrouping": "None"
				},
				{
					"AttributeName": "accountid",
					"AggregateType": "Count",
					"Alias": "cnt",
					"HasGroupBy": false,
					"DateTimeGrouping": "None"
				},
				{
					"AttributeName": "createdon",
					"AggregateType": "None",
					"Alias": "yr",
					"HasGroupBy": true,
					"DateTimeGrouping": "Year"
				},
				{
					"AttributeName": "address1_city",
					"AggregateType": "None",
					"Alias": "city",
					"HasGroupBy": true,
					"DateTimeGrouping": "None"
				}
			]
		}
	},
	"strings": {
		"Distinct": false,
		"NoLock": false,
		"QueryHints": "",
		"ForceSeek": null,
		"DataSource": null,
		"EntityName": "account",
		"TopCount": null,
		"SubQueryExpression": null,
		"PageInfo": {
			"PageNumber": 1,
			"Count": 0,
			"ReturnTotalRecordCount": false,
			"ReturnRecordCountLessThanMarv": false,
			"PagingCookie": null
		},
		"LinkEntities": [],
		"Criteria": {
			"FilterOperator": "And",
			"FilterHint": null,
			"IsQuickFindFilter": false,
			"AnyAllFilterLinkEntity": null,
			"Conditions": [
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "name",
					"Operator": "Equal",
					"Values": [
						{
							"Type": "System.String",
							"Value": "line1\nline2\ttab"
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "name",
					"Operator": "Equal",
					"Values": [
						{
							"Type": "System.String",
							"Value": "12.50"
						}
					]
				}
			],
			"Filters": []
		},
		"Orders": [],
		"ColumnSet": {
			"AllColumns": false,
			"Columns": [
				"name"
			],
			"AttributeExpressions": []
		}
	},
	"paging": {
		"Distinct": false,
		"NoLock": false,
		"QueryHints": "",
		"ForceSeek": null,
		"DataSource": null,
		"EntityName": "account",
		"TopCount": null,
		"SubQueryExpression": null,
		"PageInfo": {
			"PageNumber": 2,
			"Count": 50,
			"ReturnTotalRecordCount": true,
			"ReturnRecordCountLessThanMarv": false,
			"PagingCookie": null
		},
		"LinkEntities": [],
		"Criteria": {
			"FilterOperator": "And",
			"FilterHint": null,
			"IsQuickFindFilter": false,
			"AnyAllFilterLinkEntity": null,
			"Conditions": [],
			"Filters": []
		},
		"Orders": [],
		"ColumnSet": {
			"AllColumns": true,
			"Columns": [],
			"AttributeExpressions": []
		}
	},
	"multiSelect": {
		"Distinct": false,
		"NoLock": false,
		"QueryHints": "MaxRecursion",
		"ForceSeek": null,
		"DataSource": null,
		"EntityName": "account",
		"TopCount": null,
		"SubQueryExpression": null,
		"PageInfo": {
			"PageNumber": 1,
			"Count": 0,
			"ReturnTotalRecordCount": false,
			"ReturnRecordCountLessThanMarv": false,
			"PagingCookie": null
		},
		"LinkEntities": [],
		"Criteria": {
			"FilterOperator": "And",
			"FilterHint": null,
			"IsQuickFindFilter": false,
			"AnyAllFilterLinkEntity": null,
			"Conditions": [
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "industrycode",
					"Operator": "ContainValues",
					"Values": [
						{
							"Type": "System.Int32",
							"Value": 1
						}
					]
				},
				{
					"EntityName": null,
					"CompareColumns": false,
					"AttributeName": "parentaccountid",
					"Operator": "Under",
					"Values": [
						{
							"Type": "System.Guid",
							"Value": "00000000-0000-0000-0000-000000000001"
						}
					]
				}
			],
			"Filters": []
		},
		"Orders": [],
		"ColumnSet": {
			"AllColumns": false,
			"Columns": [
				"name"
			],
			"AttributeExpressions": []
		}
	}
} as unknown as Record<string, QeQueryExpression>;
