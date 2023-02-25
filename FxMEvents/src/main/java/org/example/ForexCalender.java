package org.example;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Locale;
import java.util.logging.SimpleFormatter;

public class ForexCalender {
    public String eventDate, time, currency, eventName;
    public String actualNumber, forcastsNumber, previousNumbers;

    DateTimeFormatter inputFormatter = DateTimeFormatter.ofPattern("MMM d yyyy", Locale.ENGLISH);
    DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");

    public ForexCalender(String eventDate, String time, String currency, String eventName, String actualNumber, String forcastsNumber, String previousNumbers) {
        LocalDate date = LocalDate.parse(eventDate, inputFormatter);
        this.eventDate = date.format(outputFormatter);
        this.time = time;
        this.currency = currency;
        this.eventName = eventName;
        this.actualNumber = actualNumber;
        this.forcastsNumber = forcastsNumber;
        this.previousNumbers = previousNumbers;
    }

    public String toString(){
        return eventDate+" "+time+" "+currency+" "+eventName+" "+actualNumber+" "+forcastsNumber+" "+previousNumbers;
    }
}
