package com.git.poitest.DAO;

import com.git.poitest.model.Student;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;
import java.util.Map;

public interface StudentRepository extends JpaRepository<Student,Integer> {


    @Query(value = "SELECT id, name, email, age FROM student", nativeQuery = true)
    List<Map<String,Object>>findAllAsMap();
}
