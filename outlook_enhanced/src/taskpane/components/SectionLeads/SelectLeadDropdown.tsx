import * as React from 'react';
import { Spinner, SpinnerSize, TextField } from 'office-ui-fabric-react';
import { _t } from '../../../utils/Translator';
import { ContentType, HttpVerb, sendHttpRequest } from '../../../utils/httpRequest';
import AppContext from '../AppContext';
import api from '../../api';
import Lead from '../../../classes/Lead';
import './SelectLeadDropdown.css';
import Logger from '../Log/Logger';
import { OdooTheme } from '../../../utils/Themes';

type SelectLeadDropdownProps = {
    onLeadClick: (lead: Lead) => void;
    canCreateLead: boolean;
    opportunityLeads: Lead[];
};

type SelectLeadDropdownState = {
    query: string;
    isLoading: boolean;
    leads: Lead[];
};

class SelectLeadDropdown extends React.Component<SelectLeadDropdownProps, SelectLeadDropdownState> {
    constructor(props, context) {
        super(props, context);
        this.state = { query: '', isLoading: false, leads: [] };
    }

    private leadsRequest;

    private onQueryChanged = (event) => {
        const query = event.target.value;
        this.setState({ query: query });
        this.cancelLeadsRequest();
        if (query.length > 0) {
            this.getLeads(query);
        } else {
            this.setState({ isLoading: false, leads: [] });
        }
    };

    private cancelLeadsRequest = () => {
        if (this.leadsRequest) this.leadsRequest.cancel();
    };

    private getLeads = async (searchTerm: string) => {
        const { opportunityLeads } = this.props;
        const existingLeadIds = opportunityLeads.map((lead) => lead.id);
        if (!searchTerm || !searchTerm.length) {
            return;
        }

        this.setState({ isLoading: true });
        this.leadsRequest = sendHttpRequest(
            HttpVerb.POST,
            api.baseURL + api.searchLead, // '/mail_plugin/lead/search'
            ContentType.Json,
            this.context.getConnectionToken(),
            { search_term: searchTerm },
            true
        );
        this.context.addRequestCanceller(this.leadsRequest.cancel);

        let response = null;
        try {
            response = JSON.parse(await this.leadsRequest.promise);
        } catch (error) {
            this.setState({ isLoading: false, leads: [] });
            this.context.showHttpErrorMessage(error);
            return;
        }

        const allLeads = response.result.map((json) => Lead.fromJSON(json));

        const filteredLeads = allLeads.filter((lead) => {
            return !existingLeadIds.includes(lead.id || lead.lead_id);
        });

        this.setState({ leads: filteredLeads, isLoading: false });
    };

    private renderLeadsList = () => {
        return (
            <div>
                {this.state.leads.map((lead) => (
                    <div key={lead.id} className="Lead-search-result-text">
                        <span onClick={() => this.props.onLeadClick(lead)}>{lead.name}</span>
                        <span className="log-email-icon">
                            <Logger
                                resId={lead.id || lead.lead_id}
                                model="crm.lead"
                                tooltipContent={`Log this email to ${lead.name}`}
                            />
                        </span>
                    </div>
                ))}

                {this.state.query.length > 0 && !this.state.leads.length && !this.state.isLoading && (
                    <div>{_t('No Lead Found')}</div>
                )}

                {this.state.isLoading && (
                    <Spinner theme={OdooTheme} size={SpinnerSize.large} className="Lead-result-spinner" />
                )}
            </div>
        );
    };

    render() {
        return (
            <div className="Lead-result-container">
                <div>{_t('Pick an Opportunity')}</div>
                <div className="Lead-search-bar">
                    <TextField
                        className="input-search"
                        placeholder={_t('Search Opportunities...')}
                        onChange={this.onQueryChanged}
                        value={this.state.query}
                        autoComplete="off"
                        onFocus={(e) => e.target.select()}
                    />
                </div>
                {this.renderLeadsList()}
            </div>
        );
    }

}

SelectLeadDropdown.contextType = AppContext;
export default SelectLeadDropdown;

