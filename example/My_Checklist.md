# Release Checklist


## Everyone

### Communication

- Check your mails for status updates
- Inform testers in advance if you require new tests to be integrated/developed *(e.g. CANoe)*
- Check the [Lessons Learned][] document

  [Lessons Learned]: https://example.com/LessonsLearned.html

### Testing Changes

- Push only tested code
- Ensure you tested also with debugger unplugged
- Review code
- Verify that your changes did not introduce new compilation warnings
- Verify that your changes did not introduce new MISRA/LINT warnings

### Publishing your changes

- Synchronize your workspace and merge your changes before pushing
- Do not push before leaving
- Check that what you pushed is compiling in a new workspace *(This ensures you published all files)*
- Verify the next day that Jenkins is still running
- Use different Change Requests for code and documentation

### Issue Tracking

- Close all your issues linked to the current delivery
- If issues are postponed to next release, inform Build Manager of possible deviation with what the customer expects
- Check that *Software Needed Action* and *Implementation Comment* are filled


## Integrators

### Preparation

- Check that we received the new part numbers from hardware
- Check production process with production team

### Integration

- Verify merge of `patch_list.txt`
- Verify new compilation warnings
- Generate MISRA/LINT report
- Run full CANoe non-regression tests
- Run dedicated integration tests for new modules
- Check for new errors: justify or fix
- Push to main branch

### Intermediary Checkpoint

- Test complete setup after mass erase with all SW parts flashed (check the [Binaries section](#binaries))
- Execute all test from integration
- Test without debugger
- Update Release Name for Jenkins
- Release CANoe simulation

### Release Checkpoint

- Close release item
- Execute all tests from intermediary checkpoint
- Execute Bootloader non-regression tests
- Do CPU load measurements
- Update Release Notes
- Gather and archive binaries (check the [Binaries section](#binaries))
- Flash and test prototypes to send to other teams


## Release Content

### Binaries

- Internal
  - `main.s19`: The application
  - `boot.s19`: The bootloader
  - `parameters.s19`: The configurable parameters
- Delivered to customer
  - `parameters.dat`: Updatable parameters

### Release Notes

- Include the detailed description of all opened issues in the release notes
- Document all known deviations
- Document all known bugs and possible workarounds
