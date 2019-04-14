package com.maybesilent.dao;

import com.maybesilent.domain.Testtable;
import com.maybesilent.domain.TesttableExample;
import java.util.List;
import org.apache.ibatis.annotations.Param;

public interface TesttableMapper {
    int countByExample(TesttableExample example);

    int deleteByExample(TesttableExample example);

    int insert(Testtable record);

    int insertSelective(Testtable record);

    List<Testtable> selectByExample(TesttableExample example);

    int updateByExampleSelective(@Param("record") Testtable record, @Param("example") TesttableExample example);

    int updateByExample(@Param("record") Testtable record, @Param("example") TesttableExample example);
}