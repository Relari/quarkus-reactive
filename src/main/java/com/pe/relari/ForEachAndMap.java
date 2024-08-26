package com.pe.relari;

import com.pe.relari.model.Employee;
import com.pe.relari.model.Person;
import com.pe.relari.repository.EmployeeRepository;
import com.pe.relari.util.EmployeeUtil;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class ForEachAndMap {

    private static final List<Employee> employees = EmployeeRepository.employees();

    private static void foreachConvertJava7FirstShape() {
        System.out.println("Foreach en Java 7");

        List<Person> people = new ArrayList<>();

        for (int i = 0; i < employees.size(); i++) {

            Employee employee = employees.get(i);
            people.add(new Person(employee.getId(), employee.getName(), employee.getSex()));

            System.out.println(employee);
        }

        EmployeeUtil.separation();

        for (int i = 0; i < people.size(); i++) {
            System.out.println(people.get(i));
        }
    }

    private static void foreachConvertJava7SecondShape() {
        System.out.println("Foreach en Java 7");

        List<Person> people = new ArrayList<>();

        for (Employee employee : employees) {
            people.add(new Person(employee.getId(), employee.getName(), employee.getSex()));

            System.out.println(employee);
        }

        EmployeeUtil.separation();

        for (Person person: people) {
            System.out.println(person);
        }
    }

    private static void foreachWithStreamAndMap() {

        List<Person> people = employees.stream()
                .map(employee -> {
                    Person person = new Person(employee.getId(), employee.getName(), employee.getSex());
                    return person;
                }).collect(Collectors.toList());

        people.forEach(System.out::println);
    }

    private static void foreachWithStreamAndMapWithBuilder() {
        List<Person> people = employees.stream()
                .map(employee -> new Person(employee.getId(), employee.getName(), employee.getSex()))
                .collect(Collectors.toList());

        people.forEach(System.out::println);
    }

    private static void foreachWithStreamAndMapWithMethod() {
        List<Person> people = employees.stream()
                .map(Person::new)
                .collect(Collectors.toList());

        people.forEach(System.out::println);
    }

    public static void main(String[] args) {

        foreachConvertJava7FirstShape();
        EmployeeUtil.separation();

        foreachConvertJava7SecondShape();
        EmployeeUtil.separation();

        foreachWithStreamAndMap();
        EmployeeUtil.separation();

        foreachWithStreamAndMapWithBuilder();
        EmployeeUtil.separation();

        foreachWithStreamAndMapWithMethod();
        EmployeeUtil.separation();

    }

}
