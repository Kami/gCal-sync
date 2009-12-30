# -*- coding: utf-8 -*-
#
# Name: Google Calendar Sync
# Description: Script for synchronizing the events matching the defined criteria
#              between two Google calendars.
# Author: Tomaž Muraus (http://www.tomaz-muraus.info)
# Version: 1.1.0
# License: GPL

# Requirements:
# - Windows / Linux / Mac OS
# - Python >= 2.6
# - Python Google Data Client (http://code.google.com/p/gdata-python-client/)

import re
import sys
import logging
import ConfigParser

import gdata.calendar.service
import gdata.service
import atom.service
import gdata.calendar

class CalendarSync:
    def __init__(self, email, password, source_calendar, target_calendar):
        """ User authentication. """
        self.calendar_service = gdata.calendar.service.CalendarService()
        self.calendar_service.email = email
        self.calendar_service.password = password
        self.calendar_service.source = 'gCal-sync-1.0'
        self.calendar_service.ProgrammaticLogin()
        
        self.source_calendar = source_calendar
        self.target_calendar = target_calendar
    
    def sync_events(self):
        """ Find the events matching the defined criteria and if they don't already exist, copy them to the target calendar.
        
        Looping over all the events and comparing them manually is needed, because Google Calendar API
        full-text query only searches the event title and content."""
        
        request_feed = gdata.calendar.CalendarEventFeed()
        matches_count = 0
        for i, event in enumerate(self.__get_future_events(self.source_calendar).entry):
            if self.__event_matches_copy_criteria(event) and not self.__event_exists(event, self.target_calendar):
                request_feed.AddInsert(entry = self.__create_event_copy(event))
                matches_count += 1
       
        logging.info('Found %d events matching the criteria (%s = %s)' % (matches_count, config.get('copy_criteria', 'field'), config.get('copy_criteria', 'value')))
        
        if matches_count > 0:
            response_feed = self.calendar_service.ExecuteBatch(request_feed, '/calendar/feeds/' + config.get('calendars', 'target_calendar') + '/private/full/batch')
            inserted = len(response_feed.entry)
            
            logging.info('%d events synchronized (%s -> %s)' % (inserted, self.source_calendar, self.target_calendar))
            
        self.__delete_orphaned_events()
            
    def __delete_orphaned_events(self):
        """ Find the future events matching the defined criteria which exist on the target but not on the source calendar
        (events which were deleted) and delete them from the target calendar."""
        
        request_feed = gdata.calendar.CalendarEventFeed()
        matches_count = 0
        for i, event in enumerate(self.__get_future_events(self.target_calendar).entry):
            if self.__event_matches_copy_criteria(event) and not self.__event_exists(event, self.source_calendar):
                request_feed.AddDelete(entry = event)
                matches_count += 1
                
        logging.info('Found %d orphaned events matching the criteria (%s = %s)' % (matches_count, config.get('copy_criteria', 'field'), config.get('copy_criteria', 'value')))
        
        if matches_count > 0:
            response_feed = self.calendar_service.ExecuteBatch(request_feed, '/calendar/feeds/' + config.get('calendars', 'target_calendar') + '/private/full/batch')
            deleted = len(response_feed.entry)
            
            logging.info('%d orphaned events deleted from the target calendar (%s)' % (deleted, self.target_calendar))

    def __get_future_events(self, calendar):
        """ Return the upcoming events.

        Keyword arguments:
        calendar -- name of the calendar
        
        """
        query = gdata.calendar.service.CalendarEventQuery(calendar, 'private', 'full')
        query.futureevents = 'true'
        query.max_results = 1000
        query.sortorder = 'ascending'
          
        events = self.calendar_service.CalendarQuery(query)
        
        return events
    
    def __event_matches_copy_criteria(self, event):
        """ Check if the event matches the defined criteria. """
        matches = False
        criteria_field = config.get('copy_criteria', 'field')
        criteria_value = config.get('copy_criteria', 'value')
        
        if criteria_field == 'title' and event.title.text != None:
            matches = True if re.search(criteria_value, event.title.text) != None else False
        elif criteria_field == 'location' and event.where[0].value_string != None:
            matches = True if re.search(criteria_value, event.where[0].value_string) != None else False
        elif criteria_field == 'content' and event.content.text != None:
            matches = True if re.search(criteria_value, event.content.text) != None else False
        
        return matches
    
    def __event_exists(self, event, calendar):
        """ Check if the same event already exists on the specified calendar. """
        query = gdata.calendar.service.CalendarEventQuery(calendar, 'private', 'full')
        query.max_results = 1000
        query.sortorder = 'ascending'
        query.start_min = event.when[0].start_time
        query.start_max = event.when[0].end_time

        events = self.calendar_service.CalendarQuery(query)

        for i, an_event in enumerate(events.entry):
            if (an_event.title.text == event.title.text and \
                an_event.where[0].value_string == event.where[0].value_string and \
                an_event.content.text == event.content.text and \
                an_event.when[0].start_time == event.when[0].start_time and \
                an_event.when[0].end_time == event.when[0].end_time):
                return True
        
        return False
    
    def __create_event_copy(self, event):
        """ Create and return a new event with the same title, start time,
        end time, location and content as the given event (basically a copy).
        
        """
        title = event.title
        when = event.when
        where = event.where
        content = event.content
        
        new_event = gdata.calendar.CalendarEventEntry()
        new_event.title = title
        new_event.content = content
        new_event.where = where
        new_event.when = when
        
        return new_event

if __name__ == '__main__':
    config = ConfigParser.ConfigParser()
    
    if not config.read('config.cfg'):
        raise Exception('Could not read the config file')

    if config.getint('other', 'enable_logging') == 1:
        logging.basicConfig(filename = config.get('other', 'log_filename'), filemode = 'a', level = logging.INFO, format = '%(asctime)s %(levelname)-8s %(message)s', datefmt = '%d.%m.%Y %H:%M:%S')
    else:
        logging.disable(logging.INFO)

    CalendarSync = CalendarSync(config.get('account_data', 'email'), config.get('account_data', 'password'), config.get('calendars', 'source_calendar'), config.get('calendars', 'target_calendar'))
    CalendarSync.sync_events()