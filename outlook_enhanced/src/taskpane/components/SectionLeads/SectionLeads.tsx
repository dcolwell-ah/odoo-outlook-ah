import * as React from 'react';
import Partner from '../../../classes/Partner';

import AppContext from '../AppContext';
import api from '../../api';
import Lead from '../../../classes/Lead';
import Section from '../Section/Section';
import SelectLeadDropdown from './SelectLeadDropdown';
import { _t } from '../../../utils/Translator';
import { Callout, DirectionalHint } from 'office-ui-fabric-react';

type LeadSectionProps = {
    partner: Partner;
    canCreatePartner: boolean;
    canCreateLead: boolean;
    opportunityLeads: Lead[];  // Add this property to the type
};

type SectionLeadsState = {
    leads: Lead[];
    isLeadSearchCalloutOpen: boolean;
};

class SectionLeads extends React.Component<LeadSectionProps, SectionLeadsState> {
    constructor(props, context) {
        super(props, context);
        this.state = { leads: this.props.partner.leads, isLeadSearchCalloutOpen: false };
    }

    private getLeadDescription = (lead: Lead): string => {
        const expectedRevenueString = _t(
            lead.recurringPlan
                ? '%(expected_revenue)s + %(recurring_revenue)s %(recurring_plan)s at %(probability)s%'
                : '%(expected_revenue)s at %(probability)s%',
            {
                expected_revenue: lead.expectedRevenue,
                recurring_revenue: lead.recurringRevenue,
                recurring_plan: lead.recurringPlan,
                probability: lead.probability,
            },
        );

        return expectedRevenueString;
    };

    private toggleLeadSearchCallout = () => {
        this.setState({ isLeadSearchCalloutOpen: !this.state.isLeadSearchCalloutOpen });
    };

    render() {
        return (
            <>
            <Section
                className="collapse-lead-section"
                records={this.state.leads}
                partner={this.props.partner}
                canCreatePartner={this.props.canCreatePartner}
                model="crm.lead"
                odooEndpointCreateRecord={api.createLead}
                odooRecordIdName="lead_id"
                odooRedirectAction="crm_mail_plugin.crm_lead_action_form_edit"
                title="Opportunities"
                titleCount="Opportunities (%(count)s)"
                msgNoPartner="Save Contact to create new Opportunities."
                msgNoPartnerNoAccess="The Contact needs to exist to create Opportunity."
                msgNoRecord="No opportunities found for this contact."
                msgLogEmail="Log Email Into Lead"
                getRecordDescription={this.getLeadDescription}
                showSearchButton={true} // <-- ONLY HERE
                onSearchButtonClick={this.toggleLeadSearchCallout}
            />
            {this.state.isLeadSearchCalloutOpen && (
                <Callout
                    directionalHint={DirectionalHint.bottomRightEdge}
                    directionalHintFixed={true}
                    onDismiss={() => this.setState({ isLeadSearchCalloutOpen: false })}
                    preventDismissOnScroll={true}
                    setInitialFocus={true}
                    doNotLayer={true}
                    gapSpace={0}
                    role="alertdialog"
                    target=".collapse-lead-section .collapse-section-button"
                >
                    <SelectLeadDropdown
                        canCreateLead={this.props.canCreateLead}
                        opportunityLeads={this.props.opportunityLeads}
                        onLeadClick={(lead) => {
                            this.setState({ isLeadSearchCalloutOpen: false });
                            // Optional: handle something with the selected lead (e.g., open detail view)
                            window.open(`${api.baseURL}/web#id=${lead.id}&model=crm.lead&view_type=form`);
                        }}
                    />
                </Callout>
            )}
        </>
        );
    }
}

SectionLeads.contextType = AppContext;

export default SectionLeads;
