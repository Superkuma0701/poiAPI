package com.git.poitest.Service.Impl;

import com.git.poitest.DAO.StudentRepository;
import com.git.poitest.Service.StudentService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;


@Service
public class StudentServiceImpl implements StudentService {


    public StudentRepository studentRepository;

    @Autowired
    public StudentServiceImpl(StudentRepository studentRepository) {
        this.studentRepository = studentRepository;
    }

    @Override
    public List<Map<String, Object>> findAll() {
        return studentRepository.findAllAsMap();
    }

    @Override
    public List<Object[]> findAlll() {
        return studentRepository.findAllByQuery();
    }
}
