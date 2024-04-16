package com.git.poitest.model;

import jakarta.persistence.*;
import lombok.Data;

@Entity
@Table(name = "student")
@Data
public class Student {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "id")
    Integer id ;
    @Column(name = "name")
    String name;
    @Column(name = "email")
    String email;
    @Column(name = "age")
    Integer age;

    public Student(String name, String email, Integer age) {
        this.name = name;
        this.email = email;
        this.age = age;
    }

    public Student() {}
}
