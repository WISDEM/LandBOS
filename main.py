import os
from datetime import datetime

from landbosse.excelio import XlsxSerialManagerRunner
from landbosse.excelio import XlsxParallelManagerRunner
from landbosse.excelio import XlsxGenerator
from landbosse.excelio import XlsxValidator
from landbosse.excelio import CsvGenerator

# LandBOSSE, small utility functions
from landbosse.excelio import XlsxFileOperations


def main():
    # Print start timestamp
    print(f'>>>>>>>> Begin run {datetime.now()} <<<<<<<<<<')

    # The file_ops object handles file names for input and output data.
    file_ops = XlsxFileOperations()

    # If run_parallel is True, an XlsxParallelManagerRunner will calculate the
    # projects in parallel. This takes advantage of multicore architecture
    # available on most hardware.
    #
    # If run_parallel is False, XlsxSerialManagerRunner will calculate projects
    # serially. This is much slower so running in parallel is preferred
    # unless there is a good reason to run serially. One such reason is using a
    # debugger which can slow down when it is being used to debug multiple
    # processes.

    run_parallel = True
    manager_runner = XlsxParallelManagerRunner(file_ops) if run_parallel else XlsxSerialManagerRunner(file_ops)

    # project_xlsx is the absolute path of the project_list.xlsx
    projects_xlsx = os.path.join(file_ops.landbosse_input_dir(), 'project_list.xlsx')

    # Should we enable the scaling study? Just look for the scaling study
    # option. The first three elements of the returned tuple aren't necessary
    # to determine if the scaling study is enabled.
    # Switch to either validation or non validation producing code.
    input_path, output_path, validation_enabled, enable_scaling_study = (
        file_ops.get_input_output_paths_from_argv_or_env()
    )

    # final_result aggregates all the results from all the projects.
    final_result = manager_runner.run_from_project_list_xlsx(projects_xlsx, enable_scaling_study)

    # Write the extended_project_list, which has all the parametric values.
    extended_project_list_path = os.path.join(file_ops.extended_project_list_path(), 'extended_project_list.csv')
    extended_project_list = final_result['extended_project_list']
    extended_project_list.to_csv(extended_project_list_path, index=False)

    # Run validation or not depending on whether validation was enabled.
    if validation_enabled:
        print('Running validation.')

        # Creates file path for output file from prior LandBOSSE run that will be used to check latest run
        # Generated based on input_path from command line when --validate option is specified
        # (validation output file must be in inputs folder and must be called 'landbosse-output-validation.xlsx')
        expected_validation_data_path = os.path.join(input_path, 'landbosse-expected-validation-data.xlsx')
        validation_result_path = os.path.join(file_ops.landbosse_output_dir(), 'landbosse-validation-result.xlsx')

        validator = XlsxValidator()
        validation_was_successful = validator.compare_expected_to_actual(
            expected_xlsx=expected_validation_data_path,
            actual_module_type_operation_list=final_result['module_type_operation_list'],
            validation_output_xlsx=validation_result_path,
        )
        if validation_was_successful:
            print('Validation passed.')
            build_status = 0
        else:
            print('Validation failed. See mismatched data above.')
            build_status = 1

    # XlsxGenerator has a context manager that writes each individual
    # worksheet to the output .xlsx. Also, copy file input structure.
    print('Writing final output folder')

    max_number_of_excel_rows = 1048576
    if len(final_result['details_list']) > max_number_of_excel_rows:
        print('WARNING: Details sheet in .xlsx has too many rows for Excel. Please use landbosse-details.csv instead.')
        print('Writing .xlsx file for backwards compatability.')

    with XlsxGenerator('landbosse-output', file_ops) as xlsx:
        xlsx.tab_costs_by_module_type_operation(rows=final_result['module_type_operation_list'])
    file_ops.copy_input_data()

    # Always write .csv versions of the output
    csv_generator = CsvGenerator(file_ops)

    costs = csv_generator.create_costs_dataframe(final_result['module_type_operation_list'])
    details = csv_generator.create_details_dataframe(final_result['details_list'])
    costs_csv_filename = os.path.join(file_ops.landbosse_output_dir(), 'landbosse-costs.csv')
    details_csv_filename = os.path.join(file_ops.landbosse_output_dir(), 'landbosse-details.csv')
    costs.to_csv(costs_csv_filename, index=False)
    details.to_csv(details_csv_filename, index=False)

    # Print end timestamp
    print(f'>>>>>>>> End run {datetime.now()} <<<<<<<<<<')

    # returns an exit code of either 0 (successful validation), or 1 (validation failed).
    #
    # If validation was not enabled, exit with a status of 0 (no errors)

    if validation_enabled:
        exit(build_status)


if __name__ == '__main__':
    main()
