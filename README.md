# kata-legacy-code

Given the codebase under excelservice-1.0, you need to improve the existing code, which can be considered as legacy code. This codebase was created ages ago, and nobody really knows who coded it or how it really works. Nonetheless, as the business changes frequently, classes has already been modified many times, so the codebase became a mess and making it hard to understand at a few places.

Today, once again, the PM wants to add some new stuff to it. As it is already quite complex, adding a new behaviour in the export is tricky. So it won't work this time to just ignore the current debt.

Your mission is to **refactor the whole codebase to make it cleaner, and thus more maintainable by the next developer**. Now is the time for you to show your exceptional skills and make this implementation better.

**The application is supposed to :**

- allow a user to create surveys
- allow a user to start a campaign based on a survey
- export an existing campaign and send it over by email

The objectives is to improve this codebase. This exercise **should not last longer than 1,5 hour**. It can be too short if you want to fix everything, but you may take add extra effort and it's up to you.

## Pre-requisites

- JDK 17

## How to run the code

You can start the application:

```
cd excelservice-1.0
./gradlew bootRun
```

This will expose a few routes [at localhost](http://localhost:8080/api). You may try this [hello world](http://localhost:8080/api/hello).

You can launch the tests with:

```
./gradlew test
```

## Deliverables

What do we expect from you:
- the link of the git repository
- several commits, with an explicit message each time
- a file / message / email explaining your process and principles you've followed

**Good luck!**
