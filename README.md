# VBA generic utilities (base functionalities extensions)
VBA is known for lacking of built-in implementations, that are common in other programming languages. Working on sophisticated real-life projects, it makes no sense (and there is no place) to repeat self-defined helper functions in each macro's procedure. That's why, in order to make your VBA-project code more efficient, clear and professional looking, it is strongly recommended to place custom extensions in subject-separeted modules and in main part of code, just refer to the function by <code> Call ModuleName.MyFunction(args) </code>. This repo includes a set of generic extensions I have written working on various projects. All of them fit to majority of projects, no matter the industry, and they can be easily customized depending on the specific needs.

There are only few assumptions to use them correctly:
    <ul>
        <li>array indexing starts with 0 - use <code>Option Base 0</code> on the top of the module code</li>
        <li>require variables declaration in advance - use <code>Option Explicit</code> on the top of the module code</li>
        <li>put data into Excel tables instead of ranges</li>

