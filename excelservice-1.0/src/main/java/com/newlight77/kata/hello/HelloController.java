package com.newlight77.kata.hello;

import org.springframework.web.bind.annotation.*;

@RestController
public class HelloController {

    @GetMapping(value = "/api/hello")
    public String createSurvey() {
        return "Hello World!";
    }

}
