/**
 * Every FetchXML example from the Microsoft Learn articles used to define FetchXML Studio's
 * behaviour (#90): full queries, plus bare fragments wrapped in a minimal <fetch>.
 *
 * Source: Microsoft Learn (MicrosoftDocs/powerapps-docs). Code samples © Microsoft Corporation,
 * licensed under the MIT License: https://github.com/MicrosoftDocs/powerapps-docs/blob/main/LICENSE-CODE
 *
 * Generated — regenerate rather than hand-edit.
 */

export interface LearnFetchXmlExample {
	/** Learn article the example comes from */
	source: string;
	xml: string;
}

export const LEARN_FETCHXML_EXAMPLES: LearnFetchXmlExample[] = [
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/aggregate-data",
		"xml": "<fetch aggregate='true'>\n  <entity name='account'>\n    <attribute name='numberofemployees'\n      alias='Average'\n      aggregate='avg' />\n    <attribute name='numberofemployees'\n      alias='Count'\n      aggregate='count' />\n    <attribute name='numberofemployees'\n      alias='ColumnCount'\n      aggregate='countcolumn' />\n    <attribute name='numberofemployees'\n      alias='Maximum'\n      aggregate='max' />\n    <attribute name='numberofemployees'\n      alias='Minimum'\n      aggregate='min' />\n    <attribute name='numberofemployees'\n      alias='Sum'\n      aggregate='sum' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/aggregate-data",
		"xml": "<fetch aggregate='true'>\n   <entity name='account'>\n      <attribute name='numberofemployees'\n         alias='Total'\n         aggregate='sum' />\n      <attribute name='address1_city'\n         alias='Count'\n         aggregate='count' />\n      <attribute name='address1_city'\n         alias='City'\n         groupby='true' />\n      <order alias='City' />\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/aggregate-data",
		"xml": "<fetch aggregate='true'>\n   <entity name='account'>\n      <attribute name='numberofemployees'\n         alias='Total'\n         aggregate='sum' />\n      <attribute name='createdon'\n         alias='Day'\n         groupby='true'\n         dategrouping='day' />\n      <attribute name='createdon'\n         alias='Week'\n         groupby='true'\n         dategrouping='week' />\n      <attribute name='createdon'\n         alias='Month'\n         groupby='true'\n         dategrouping='month' />\n      <attribute name='createdon'\n         alias='Year'\n         groupby='true'\n         dategrouping='year' />\n      <attribute name='createdon'\n         alias='FiscalPeriod'\n         groupby='true'\n         dategrouping='fiscal-period' />\n      <attribute name='createdon'\n         alias='FiscalYear'\n         groupby='true'\n         dategrouping='fiscal-year' />\n      <order alias='Month' />\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/aggregate-data",
		"xml": "<fetch aggregate=\"true\">\n   <entity name=\"order\">\n      <attribute name=\"totalamount\"\n         aggregate=\"sum\"\n         alias=\"total\" />\n      <attribute name=\"datefulfilled\"\n         groupby=\"true\"\n         dategrouping=\"fiscal-period\" />\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/aggregate-data",
		"xml": "<fetch top='5'>\n   <entity name='account'>\n      <attribute name='name' />\n      <attribute name='accountid'\n         alias='numberOfChildren'\n         rowaggregate='CountChildren' />\n      <order attribute='accountid'\n         descending='true' />\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/aggregate-data",
		"xml": "<fetch aggregate='true'\n   aggregatelimit = '5'>\n   <entity name='account'>\n      <attribute name='name'\n         alias='account_count'\n         aggregate='count' />\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/attribute",
		"xml": "<fetch>\n  <entity name='account'>\n    <attribute name='name' />\n    <link-entity name='contact'\n      from='contactid'\n      to='primarycontactid'\n      link-type='inner'\n      alias='contact'>\n      <attribute name='fullname' />\n    </link-entity>\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/condition",
		"xml": "<fetch>\n  <entity name='account'>\n    <attribute name='name' />\n    <filter type='and'>\n      <condition attribute='address1_city'\n        operator='eq'\n        value='Redmond' />\n    </filter>\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/entity",
		"xml": "<fetch>\n  <entity name='account'>\n    <attribute name='accountclassificationcode' />\n    <attribute name='createdby' />\n    <attribute name='createdon' />\n    <attribute name='name' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/fetch",
		"xml": "<fetch>\n  <entity name='account' />\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows",
		"xml": "<fetch>\n   <entity name='account'>\n      <attribute name='name' />\n      <attribute name='address1_city' />\n      <filter type='or'>\n         <condition attribute='address1_city'\n            operator='eq'\n            value='Redmond' />\n         <condition attribute='address1_city'\n            operator='eq'\n            value='Seattle' />\n         <condition attribute='address1_city'\n            operator='eq'\n            value='Bellevue' />\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows",
		"xml": "<fetch>\n   <entity name='account'>\n      <attribute name='name' />\n      <attribute name='address1_city' />\n      <filter type='and'>\n         <condition attribute='address1_city'\n            operator='in'>\n            <value>Redmond</value>\n            <value>Seattle</value>\n            <value>Bellevue</value>\n         </condition>\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows",
		"xml": "<fetch>\n  <entity name='contact'>\n    <attribute name='fullname' />\n    <filter>\n      <condition entityname='a'\n        attribute='fax'\n        operator='null' />\n    </filter>\n    <link-entity name='account'\n      from='accountid'\n      to='parentcustomerid'\n      link-type='outer'\n      alias='a' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows",
		"xml": "<fetch>\n   <entity name='contact' >\n      <attribute name='firstname' />\n      <filter>\n         <condition attribute='firstname'\n            operator='eq'\n            valueof='lastname' />\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows",
		"xml": "<fetch>\n   <entity name='contact'>\n      <attribute name='contactid' />\n      <attribute name='fullname' />\n      <filter type='and'>\n         <condition attribute='fullname'\n            operator='eq'\n            valueof='acct.name' />\n      </filter>\n      <link-entity name='account'\n         from='accountid'\n         to='parentcustomerid'\n         link-type='outer'\n         alias='acct'>\n         <attribute name='name' />\n      </link-entity>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows",
		"xml": "<fetch>\n   <entity name='contact'>\n      <attribute name='fullname' />\n      <filter type='or'>\n         <link-entity name='account'\n            from='primarycontactid'\n            to='contactid'\n            link-type='any'>\n            <filter type='and'>\n               <condition attribute='name'\n                  operator='eq'\n                  value='Contoso' />\n            </filter>\n         </link-entity>\n         <condition attribute='statecode'\n            operator='eq'\n            value='1' />\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows",
		"xml": "<fetch>\n   <entity name='contact'>\n      <attribute name='fullname' />\n      <filter type='and'>\n         <link-entity name='account'\n            from='primarycontactid'\n            to='contactid'\n            link-type='not any'>\n            <filter type='and'>\n               <condition attribute='name'\n                  operator='eq'\n                  value='Contoso' />\n            </filter>\n         </link-entity>\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows",
		"xml": "<fetch>\n   <entity name='contact'>\n      <attribute name='fullname' />\n      <filter type='and'>\n         <link-entity name='account'\n            from='primarycontactid'\n            to='contactid'\n            link-type='not all'>\n            <filter type='and'>\n               <condition attribute='name'\n                  operator='eq'\n                  value='Contoso' />\n            </filter>\n         </link-entity>\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows",
		"xml": "<fetch>\n   <entity name='contact'>\n      <attribute name='fullname' />\n      <filter type='and'>\n         <link-entity name='account'\n            from='primarycontactid'\n            to='contactid'\n            link-type='all'>\n            <filter type='and'>\n               <condition attribute='name'\n                  operator='eq'\n                  value='Contoso' />\n            </filter>\n         </link-entity>\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables",
		"xml": "<fetch top='5'>\n   <entity name='account'>\n      <attribute name='name' />\n      <link-entity name='contact'\n         from='contactid'\n         to='primarycontactid'\n         link-type='inner'\n         alias='contact'>\n         <attribute name='fullname' />\n      </link-entity>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables",
		"xml": "<fetch top='5'>\n   <entity name='contact'>\n      <attribute name='fullname' />\n      <link-entity name='account'\n         from='primarycontactid'\n         to='contactid'\n         alias='account'>\n         <attribute name='name' />\n      </link-entity>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables",
		"xml": "<fetch top='2'>\n   <entity name='systemuser'>\n      <attribute name='fullname' />\n      <link-entity name='teammembership'\n         from='systemuserid'\n         to='systemuserid'\n         intersect='true' >\n         <link-entity name='team'\n            from='teamid'\n            to='teamid'\n            link-type='inner'\n            alias='team'>\n            <attribute name='name' />\n         </link-entity>\n      </link-entity>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables",
		"xml": "<fetch>\n   <entity name='account'>\n     <attribute name='name' />\n     <link-entity name='contact'\n       from='fullname'\n       to='name'\n       link-type='inner'\n       alias='contact'>\n       <attribute name='fullname' />\n     </link-entity>\n   </entity>\n </fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables",
		"xml": "<fetch>\n   <entity name='account'>\n      <attribute name='name' />\n      <order attribute='name' />\n      <link-entity name='contact'\n         from='parentcustomerid'\n         to='accountid'\n         link-type='outer'\n         alias='contact' />\n      <filter type='and'>\n         <condition entityname='contact'\n            attribute='parentcustomerid'\n            operator='null' />\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables",
		"xml": "<fetch>\n   <entity name='contact'>\n      <attribute name='fullname' />\n      <link-entity name='account'\n         from='primarycontactid'\n         to='contactid'\n         link-type='exists'>\n         <filter type='and'>\n            <condition attribute='statecode'\n               operator='eq'\n               value='1' />\n         </filter>\n      </link-entity>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables",
		"xml": "<fetch>\n   <entity name='contact'>\n      <attribute name='fullname' />\n      <link-entity name='account'\n         from='primarycontactid'\n         to='contactid'\n         link-type='in'>\n         <filter type='and'>\n            <condition attribute='statecode'\n               operator='eq'\n               value='1' />\n         </filter>\n      </link-entity>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables",
		"xml": "<fetch>\n   <entity name='contact'>\n      <attribute name='fullname' />\n      <link-entity name='account'\n         from='primarycontactid'\n         to='contactid'\n         link-type='matchfirstrowusingcrossapply'>\n         <attribute name='accountid' />\n         <attribute name='name' />\n      </link-entity>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/link-entity",
		"xml": "<fetch>\n  <entity name='contact'>\n    <attribute name='fullname' />\n    <link-entity name='account' \n     from='primarycontactid' \n     to='contactid' \n     alias='account'>\n      <attribute name='name' />\n    </link-entity>\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/link-entity",
		"xml": "<fetch>\n  <entity name='systemuser'>\n    <attribute name='fullname' />\n    <link-entity name='teammembership'\n      from='systemuserid'\n      to='systemuserid' >\n      <link-entity name='team'\n        from='teamid'\n        to='teamid'\n        link-type='inner'\n        alias='team'>\n        <attribute name='name' />\n      </link-entity>\n    </link-entity>\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/link-entity",
		"xml": "<fetch top=\"2\">\n         <entity name=\"systemuser\">\n            <attribute name=\"fullname\" />\n            <link-entity\n               name=\"teammembership\"\n               from=\"systemuserid\"\n               to=\"systemuserid\"\n               intersect=\"true\"\n            >\n               <link-entity\n               name=\"team\"\n               from=\"teamid\"\n               to=\"teamid\"\n               link-type=\"inner\"\n               alias=\"team\"\n               >\n               <attribute name=\"name\" />\n               </link-entity>\n            </link-entity>\n         </entity>\n      </fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/link-entity",
		"xml": "<fetch top=\"2\">\n         <entity name=\"systemuser\">\n            <attribute name=\"fullname\" />\n            <link-entity name=\"team\" alias=\"team\">\n               <attribute name=\"name\" />\n            </link-entity>\n         </entity>\n      </fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators",
		"xml": "<fetch>\n   <entity name=\"order\">\n      <attribute name=\"name\" />\n      <filter type=\"and\">\n         <condition attribute=\"datefulfilled\"\n            operator=\"in-fiscal-period\"\n            value=\"3\" />\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators",
		"xml": "<fetch>\n   <entity name=\"order\">\n      <attribute name=\"name\" />\n      <filter type=\"and\">\n         <condition attribute=\"datefulfilled\"\n            operator=\"in-fiscal-period-and-year\">\n            <value>3</value>\n            <value>2023</value>\n         </condition>\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators",
		"xml": "<fetch>\n   <entity name=\"account\">\n      <attribute name=\"name\" />\n      <filter type=\"and\">\n         <condition attribute=\"createdon\"\n            operator=\"in-fiscal-year\"\n            value=\"2023\" />\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators",
		"xml": "<fetch>\n   <entity name=\"order\">\n      <attribute name=\"name\" />\n      <filter type=\"and\">\n         <condition attribute=\"datefulfilled\"\n            operator=\"last-fiscal-period\" />\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators",
		"xml": "<fetch>\n   <entity name=\"opportunity\">\n      <attribute name=\"name\" />\n      <filter type=\"and\">\n         <condition attribute=\"estimatedclosedate\"\n            operator=\"next-x-fiscal-years\"\n            value=\"3\" />\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/operators",
		"xml": "<fetch>\n   <entity name=\"incident\">\n      <attribute name=\"title\" />\n      <attribute name=\"ticketnumber\" />\n      <attribute name=\"createdon\" />\n      <attribute name=\"incidentid\" />\n      <filter type=\"and\">\n         <condition attribute=\"createdon\"\n            operator=\"olderthan-x-minutes\"\n            value=\"30\" />\n      </filter>\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/optimize-performance",
		"xml": "<fetch>\n   <entity name=\"email\">\n      <attribute name=\"activityid\" />\n      <attribute name=\"subject\" />\n      <filter type=\"and\">\n         <condition attribute=\"subject\"\n            operator=\"like\"\n            value=\"Alert:%\" />\n         <condition attribute=\"statecode\"\n            operator=\"eq\"\n            value=\"0\" />\n         <filter type=\"or\"\n            hint=\"union\">\n            <condition attribute=\"telephone1\"\n               operator=\"eq\"\n               value=\"555-123-4567\"\n               entityname=\"ac\" />\n            <condition attribute=\"telephone1\"\n               operator=\"eq\"\n               value=\"555-123-4567\"\n               entityname=\"co\" />\n         </filter>\n      </filter>\n      <link-entity name=\"account\"\n         from=\"accountid\"\n         to=\"regardingobjectid\"\n         link-type=\"outer\"\n         alias=\"ac\" />\n      <link-entity name=\"contact\"\n         from=\"contactid\"\n         to=\"regardingobjectid\"\n         link-type=\"outer\"\n         alias=\"co\" />\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/order-rows",
		"xml": "<fetch>\n  <entity name='account'>\n    <attribute name='name' />\n    <attribute name='accountnumber' />\n    <attribute name='createdon' />\n    <order attribute='createdon' />\n    <order attribute='name' />\n    <order attribute='accountnumber' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/order-rows",
		"xml": "<fetch>\n  <entity name='account'>\n    <attribute name='name' />\n    <attribute name='createdon' />\n    <order attribute='createdon' descending='true' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/order-rows",
		"xml": "<fetch>\n  <entity name='account'>\n    <attribute name='name' />\n    <attribute name='accountnumber' />\n    <attribute name='createdon' />\n    <link-entity name='account'\n      from='accountid'\n      to='parentaccountid'\n      link-type='inner'\n      alias='parentaccount'>\n      <attribute name='name'\n        alias='parentaccount' />\n        <!-- The link-entity parentaccount name -->\n      <order attribute='name' />\n    </link-entity>\n    <!-- The entity account name -->\n    <order attribute='name' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/order-rows",
		"xml": "<fetch>\n  <entity name='account'>\n    <attribute name='name' />\n    <attribute name='accountnumber' />\n    <attribute name='createdon' />\n    <link-entity name='account'\n      from='accountid'\n      to='parentaccountid'\n      link-type='inner'\n      alias='parentaccount'>\n      <attribute name='name'\n        alias='parentaccount' />\n    </link-entity>\n    <!-- The link-entity parentaccount name -->\n    <order entityname='parentaccount'\n      attribute='name' />\n      <!-- The entity account name -->\n    <order attribute='name' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/overview",
		"xml": "<fetch top='5'>\n  <entity name='account'>\n      <attribute name='name' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results",
		"xml": "<fetch count='3' page='1'>\n  <entity name='account'>\n    <attribute name='name' />\n    <order attribute='name' />\n    <order attribute='accountid' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results",
		"xml": "<fetch count='3' page='2'>\n  <entity name='account'>\n    <attribute name='name' />\n    <order attribute='name' />\n    <order attribute='accountid' />    \n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results",
		"xml": "<fetch count='3' page='1'>\n                <entity name='contact'>\n                    <attribute name='fullname'/>\n                    <attribute name='jobtitle'/>\n                    <attribute name='annualincome'/>\n                    <order descending='true' attribute='fullname'/>\n                </entity>\n        </fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results",
		"xml": "<fetch count='3' page='1'>\n  <entity name='contact'>\n    <attribute name='fullname' />\n    <attribute name='jobtitle' />\n    <attribute name='annualincome' />\n    <order descending='true' attribute='fullname' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results",
		"xml": "<fetch count='3' page='2' paging-cookie='&lt;cookie page=\"1\"&gt;&lt;fullname last=\"Susanna Stubberod (sample)\" first=\"Yvonne McKay (sample)\" /&gt;&lt;contactid last=\"{70BF4D48-34CB-ED11-B596-0022481D68CD}\" first=\"{49B0BE2E-D01C-ED11-B83E-000D3A572421}\" /&gt;&lt;/cookie&gt;'>\n  <entity name='contact'>\n    <attribute name='fullname' />\n    <attribute name='jobtitle' />\n    <attribute name='annualincome' />\n    <order descending='true' attribute='fullname' />\n  </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results",
		"xml": "<fetch>\n      <entity name='contact'>\n         <attribute name='fullname'/>\n         <attribute name='jobtitle'/>\n         <attribute name='annualincome'/>\n         <order descending='true' attribute='fullname'/>\n      </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/query-hierarchical-data",
		"xml": "<fetch>  \n  <entity name='account'>  \n    <attribute name='name' />  \n    <attribute name='accountid' />  \n    <attribute name='accountid' rowaggregate='CountChildren' alias='AccountChildren'/>  \n    <filter type='and'>  \n      <condition attribute='accountid' operator='under' value='{0}' />  \n    </filter>  \n  </entity>  \n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/quick-find",
		"xml": "<fetch>\n           <entity name='account'>\n              <attribute name='accountid' />\n              <attribute name='name' />\n              <attribute name='accountnumber' />\n              <attribute name='primarycontactid' />\n              <attribute name='address1_city' />\n              <attribute name='telephone1' />\n              <attribute name='emailaddress1' />\n              <filter type='or'\n                 isquickfindfields='1'>\n                 <condition attribute='telephone2'\n                    operator='like'\n                    value='{0}' />\n                 <condition attribute='telephone1'\n                    operator='like'\n                    value='{0}' />\n                 <condition attribute='emailaddress1'\n                    operator='like'\n                    value='{0}' />\n                 <condition attribute='accountnumber'\n                    operator='like'\n                    value='{0}' />\n                 <condition attribute='name'\n                    operator='like'\n                    value='{0}' />\n              </filter>\n              <filter type='and'>\n                 <condition attribute='statecode'\n                    operator='eq'\n                    value='0' />\n              </filter>\n              <order attribute='name'\n              descending='false' />\n           </entity>\n        </fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/select-columns",
		"xml": "<fetch top='3'>\n            <entity name='account'>\n              <attribute name='accountclassificationcode' alias='code' />\n              <attribute name='createdby' alias='whocreated' />\n              <attribute name='createdon' alias='whencreated' />\n              <attribute name='name' alias='companyname' />\n            </entity>\n          </fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/select-columns",
		"xml": "<fetch>\n   <entity name='account'>\n      <attribute name='accountclassificationcode' alias='code' />\n      <attribute name='createdby' alias='whocreated' />\n      <attribute name='createdon' alias='whencreated' />\n      <attribute name='name' alias='companyname' />\n   </entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/aggregate-data (fragment)",
		"xml": "<fetch>\n<entity name=\"account\">\n<attribute name='numberofemployees' \n   alias='ColumnCount' \n   aggregate='countcolumn' \n   distinct='true' />\n</entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows (fragment)",
		"xml": "<fetch>\n<entity name=\"account\">\n<filter><condition attribute='ownerid'\n   operator='eq-userid' /></filter>\n</entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows (fragment)",
		"xml": "<fetch>\n<entity name=\"account\">\n<filter><condition attribute='statecode'\n   operator='eq'\n   value='0' /></filter>\n</entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows (fragment)",
		"xml": "<fetch>\n<entity name=\"account\">\n<filter><condition attribute=\"numberofemployees\"\n   operator=\"between\">\n   <value>6</value>\n   <value>20</value>\n</condition></filter>\n</entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/order-rows (fragment)",
		"xml": "<fetch>\n<entity name=\"account\">\n<order attribute='accountnumber' />   \n<order attribute='createdon' />\n<order attribute='name' />\n</entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/query-antipatterns (fragment)",
		"xml": "<fetch>\n<entity name=\"account\">\n<filter><condition attribute='accountnumber'\n   operator='like'\n   value='%234' /></filter>\n</entity>\n</fetch>"
	},
	{
		"source": "https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/value (fragment)",
		"xml": "<fetch>\n<entity name=\"account\">\n<filter>\n   <condition attribute=\"numberofemployees\" operator=\"between\">\n      <value>6</value>\n      <value>20</value>\n   </condition>\n</filter>\n</entity>\n</fetch>"
	}
];
