module TimeFinder.Models {

    export interface IEmailAddress {
        address?: string;
        name?:string;
    }
    export interface IAttendee {
        type?: string;
        emailAddress?: IEmailAddress;
    }

    export interface IMeetingTimesRequest {
        attendees?: IAttendee[],
        isOrganizerOptional?:boolean,
        locationConstraint?: any,
        maxCandidates?: number,
        meetingDuration?: string;
        minimumAttendeePercentage?: number,
        returnSuggestionReasons?: boolean,
        timeConstraint?:any;
    }

    export interface IMeetingTimeSuggestion  {
        attendeeAvailability?: any;
        confidence?: number;
        locations?: any;
        meetingTimeSlot?: any;
        organizerAvailability?: any;
        suggestionReason?: number;
    }

    export interface IMeetingTimeSuggestionsResult  {
        meetingTimeSuggestions?: IMeetingTimeSuggestion[];
        emptySuggestionsReason?: any;
}

    export interface IUser {
        displayName?: string;
        givenName?: string;
        id?: string;
        jobTitle?: string;
        mail?: string;
        mobilePhone?: number;
        officeLocation?: string;
        preferredLanguage?: string;
        surname?: string;
        userPrincipalName?: string;
    }
}