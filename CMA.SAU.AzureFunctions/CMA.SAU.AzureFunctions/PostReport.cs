using Azure.Storage.Blobs.Models;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.News.DataModel;
using Portable.Xaml.Markup;
using System;
using System.Collections.Generic;
using System.Net;
using static System.Net.Mime.MediaTypeNames;

namespace CMA.SAU.AzureFunctions
{
    internal class PostReport
    {
        internal static void Submit(ILogger log, Response response, dynamic payload)
        {
            string submission_list = System.Environment.GetEnvironmentVariable("POST_REPORT_LIST");
            using (ClientContext ctx = Utilities.GetSAUCasesContext())
            {
                string[] pa_names = RemoveBlankEntries(payload.post_report.pa_names.ToObject<string[]>());
                FieldLookupValue[] names = UpdatePANames(ctx, pa_names);
                List list = ctx.Web.Lists.GetByTitle(submission_list);

                ListItem listItem = GetExistingRecord(list, (string)payload.reference);
                if (listItem == null)
                {
                    ListItemCreationInformation lici = new ListItemCreationInformation();
                    listItem = list.AddItem(lici);
                }

                listItem[Constants.TITLE] = ((string)payload.reference);
                listItem["SAUSchemeSubsidy"] = ((string)payload.scheme_subsidy);
                listItem["SAUReferralType"] = ((string)payload.referral_type);
                listItem["SAUIsC2Relevant"] = ((string)payload.is_c2_relevant) == "y" ? "Yes" : "No";
                listItem["SAUIsP3Relevant"] = ((string)payload.is_p3_relevant) == "y" ? "Yes" : "No";
                listItem["SAUEEAssessReq"] = ((string)payload.ee_assess_required) == "y" ? "Yes" : "No";
                listItem["SAUSubsidyForms"] = payload.subsidy_forms.ToObject<string[]>();
                listItem["SAUSectors"] = payload.sectors?.ToObject<string[]>();
                listItem["SAUPurposes"] = payload.purposes.ToObject<string[]>();
                listItem["SAULocations"] = payload.locations.ToObject<string[]>();
                listItem["SAUBeneficiary"] = ((string)payload.beneficiary);
                listItem["SAUBenSize"] = ((string)payload.ben_size);
                listItem["SAUGoodsServices"] = payload.ben_good_svr.ToObject<string[]>();
                listItem["SAUSpecialCatValues"] = payload.special_cat_values.ToObject<string[]>();
                listItem["SAUStartDate"] = payload.start_date != null ? ((DateTime)payload.start_date) : null;
                listItem["SAUEndDate"] = payload.end_date != null ? ((DateTime)payload.end_date) : null;
                listItem["SAUSubmittedDate"] = payload.submitted_date != null ? ((DateTime)payload.submitted_date) : null;
                listItem["SAUCompletedDate"] = payload.completed_date != null ? ((DateTime)payload.completed_date) : null;
                listItem["SAUPANames"] = names;

                dynamic postReport = payload.post_report;

                listItem["SAUReferralName"] = (string)postReport.referral_name;
                listItem["SAUPEPolicy"] = SetYesNoChoiceField(postReport.pe_policy);
                listItem["SAUPEOtherMeans"] = SetYesNoChoiceField(postReport.pe_other_means);
                listItem["SAUCCounterfactual"] = SetYesNoChoiceField(postReport.pc_counterfactual);
                listItem["SAUCEcoBehaviour"] = SetYesNoChoiceField(postReport.pc_eco_behaviour);
                listItem["SAUDAdditionality"] = SetYesNoChoiceField(postReport.pd_additionality);
                listItem["SAUDCosts"] = SetYesNoChoiceField(postReport.pd_costs);
                listItem["SAUBProportion"] = SetYesNoChoiceField(postReport.pb_proportion);
                listItem["SAUFSubsidyChars"] = SetYesNoChoiceField(postReport.pf_subsidy_char);
                listItem["SAUFMarketChars"] = SetYesNoChoiceField(postReport.pf_market_char);
                listItem["SAUGBalanceUK"] = SetYesNoChoiceField(postReport.pg_balance_uk);
                listItem["SAUGBalanceIntl"] = SetYesNoChoiceField(postReport.pg_balance_intl);

                listItem["SAUAPolicyEvidence"] = SetYesNoChoiceField(postReport.pa_policy_evidence);
                listItem["SAUAMarketFail"] = SetYesNoChoiceField(postReport.pa_market_fail);
                listItem["SAUAEquity"] = SetYesNoChoiceField(postReport.pa_equity);

                listItem["SAUEEPrinciples"] = SetYesNoChoiceField(postReport.ee_principles);
                listItem["SAUEEIssues"] = SetYesNoChoiceField(postReport.ee_issues);
                listItem["SAUOtherIssues"] = SetYesNoChoiceField(postReport.other_issues);

                listItem["SAUSpecialCats"] = SetYesNoChoiceField(postReport.special_cats);
                listItem["SAUThirdPartyReps"] = SetYesNoChoiceField(postReport.third_party_reps);
                listItem["SAUConfiIssues"] = SetYesNoChoiceField(postReport.confi_issues);

                listItem["SAUEERequired"] = SetYesNoChoiceField(postReport.ee_required);

                listItem["SAUValue"] = (string)postReport.value;
                // listItem["SAUOtherIssueLink"] = postReport.other_issues_link
                listItem["SAURejectReason"] = (string)postReport.reject_reason;
                listItem["SAUWithdrawnReason"] = (string)postReport.withdrawn_reason;

                SetTextField(listItem, "SAUPEPolicy", (string)postReport.pe_policy_text);
                SetTextField(listItem, "SAUPEOtherMeans", (string)postReport.pe_other_means_text);
                SetTextField(listItem, "SAUCCounterfactual", (string)postReport.pc_counterfactual_text);
                SetTextField(listItem, "SAUCEcoBehaviour", (string)postReport.pc_eco_behaviour_text);
                SetTextField(listItem, "SAUDAdditionality", (string)postReport.pd_additionality_text);
                SetTextField(listItem, "SAUDCosts", (string)postReport.pd_costs_text);
                SetTextField(listItem, "SAUBProportion", (string)postReport.pb_proportion_text);
                SetTextField(listItem, "SAUFSubsidyChars", (string)postReport.pf_subsidy_char_text);
                SetTextField(listItem, "SAUFMarketChars", (string)postReport.pf_market_char_text);
                SetTextField(listItem, "SAUGBalanceUK", (string)postReport.pg_balance_uk_text);
                SetTextField(listItem, "SAUGBalanceIntl", (string)postReport.pg_balance_intl_text);
                SetTextField(listItem, "SAUAPolicyEvidence", (string)postReport.pa_policy_evidence_text);
                SetTextField(listItem, "SAUAMarketFail", (string)postReport.pa_market_fail_text);
                SetTextField(listItem, "SAUAEquity", (string)postReport.pa_equity_text);
                SetTextField(listItem, "SAUEEPrinciples", (string)postReport.ee_principles_text);
                SetTextField(listItem, "SAUEEIssues", (string)postReport.ee_issues_text);
                SetTextField(listItem, "SAUOtherIssues", (string)postReport.other_issues_text);
                SetTextField(listItem, "SAUThirdPartyReps", (string)postReport.third_party_reps_text);
                SetTextField(listItem, "SAUConfiIssues", (string)postReport.confi_issues_text);

                listItem.Update();
                ctx.ExecuteQueryRetry();
            }
        }

        private static void SetTextField(ListItem listItem, string fieldName, string text)
        {
            string textField = $"{fieldName}Text";
            if ((string)listItem[fieldName] == "Yes")
            {
                listItem[textField] = text;
            }
            else
            {
                listItem[textField] = "";
            }
        }

        private static string SetYesNoChoiceField(dynamic pe_policy)
        {
            if (!string.IsNullOrEmpty((string)pe_policy))
            {
                return ((string)pe_policy) == "y" ? "Yes" : "No";
            }
            return null;
        }

        private static ListItem GetExistingRecord(List list, string reference)
        {
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            query.ViewXml = "<View Scope='RecursiveAll'><Query>" +
                            $"<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>{reference}</Value></Eq></Where>" +
                            "</Query><ViewFields><FieldRef Name='Title'/></ViewFields></View>";

            ListItemCollection items = list.GetItems(query);
            list.Context.Load(items);
            list.Context.ExecuteQueryRetry();

            return items.Count > 0 ? items[0] : null;
        }

        private static string[] RemoveBlankEntries(string[] array)
        {
            List<string> temp = new();
            foreach (string s in array)
            {
                if (!string.IsNullOrEmpty(s)) temp.Add(s);
            }
            return temp.ToArray();
        }

        private static FieldLookupValue[] UpdatePANames(ClientContext ctx, string[] pa_names)
        {
            List<FieldLookupValue> ids = new();
            string lookupList = System.Environment.GetEnvironmentVariable("PA_NAMES_LOOKUP_LIST");
            List list = ctx.Web.Lists.GetByTitle(lookupList);
            foreach (string item in pa_names)
            {
                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                query.ViewXml = "<View Scope='RecursiveAll'><Query>" +
                                $"<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>{item}</Value></Eq></Where>" +
                                "</Query><ViewFields><FieldRef Name='Title'/></ViewFields></View>";

                ListItemCollection items = list.GetItems(query);
                list.Context.Load(items);
                list.Context.ExecuteQueryRetry();

                if (items.Count == 0)
                {
                    // Add new item to list
                    ListItemCreationInformation lici = new ListItemCreationInformation();
                    ListItem listItem = list.AddItem(lici);
                    listItem[Constants.TITLE] = item;
                    listItem.Update();
                    //list.Context.Load(listItem);
                    list.Context.ExecuteQueryRetry();

                    ids.Add(new FieldLookupValue() { LookupId = listItem.Id });
                }
                else
                {
                    ids.Add(new FieldLookupValue() { LookupId = items[0].Id });
                }
            }
            return ids.ToArray();
        }
    }
}