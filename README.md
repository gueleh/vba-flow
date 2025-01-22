# vba-flow

## Introduction

VBA.Flow() is the name of my way of working with Excel VBA, an approach that evolved over many years. You can learn more about this here: https://vba-flow.net/.

This repository contains a library of solutions based on Excel VBA.

It is being built gradually.

Unless otherwise specified, each solution is contained in a separate folder.

These solutions either work independently or are based on Flow Framework 2, please refer to https://github.com/gueleh/flowframework2 to learn more about this framework.

Unless otherwise specified, you can use these solutions freely as long as you do not use them commercially. In case of commercial usage please contact me and we'll reach an agreement for your usage.

## Solutions

### Basic Semi-Automated Testing
With this small solution you can easily write and run tests when working with Excel VBA. It is not a test suite, it is not fully automated, but it is small and easy to use, with the potential to add a lot of value to your development experience. You can find this solution in the subfolder https://github.com/gueleh/vba-flow/tree/main/basic-semi-automated-testing.

### SharePoint Drive Mapper
One small class to temporarily mount a drive based on a network path without credentials, i.e. the user running the code must have access rights via the machine he runs the code on. Used to allow for file and folder processing with the standard file system. You can this this solution in the subfolder https://github.com/gueleh/vba-flow/tree/main/sharepoint-drive-mapper.

### User Defined Functions
A package with some user defined functions, i.e. functions that can be used in cells. Also includes a class for registering a user defined function in the function wizard of Excel and applies this to the functions of the package. You can find this solution in the subfolder https://github.com/gueleh/vba-flow/tree/main/user-defined-functions.

### Version Control Data Generator
With this solution you can easily generate files for using version control with Excel. This includes the code modules, but also other important data like named ranges, worksheet meta data, defined worksheet contents and contents of settings sheets, the two worksheet-related bits by contract, i.e. the structure must meet certain criteria. You can find this solution in the subfolder https://github.com/gueleh/vba-flow/tree/main/version-control-data-generator.

### CRC32 Native Hashing
Calculate CRC32 hash values without needing external references. Please read the following remarks before using this.

#### **Limits of CRC32 Hashing**

The limitations of CRC32 make it unsuitable for security-critical or demanding hashing applications. However, it is a good choice for simple error detection and checksum use cases. For scenarios requiring security or collision resistance, cryptographic hash functions like SHA-256, SHA-3, or BLAKE3 should be used.

##### 1. **Not Cryptographically Secure**
- **Weakness**: CRC32 is not designed for ensuring data integrity against intentional manipulation.
- **Reason**: It is vulnerable to deliberate collisions because its structure can be easily analyzed.
- **Implication**: It should not be used for passwords, digital signatures, or other security-critical applications.

##### 2. **Collisions**
- **Weakness**: CRC32 has only a 32-bit output (4 bytes), meaning there are only \( 2^{32} \) (around 4.3 billion) possible values.
- **Reason**: With large datasets or similar inputs, collisions (two different inputs producing the same hash) are highly likely.
- **Implication**: CRC32 is unsuitable for applications requiring many unique hashes.

##### 3. **Weak for Similar Inputs**
- **Weakness**: Small changes in the input string don’t always produce drastically different hash values.
- **Reason**: CRC32 is a linear algorithm and does not offer complex diffusion of input data.
- **Implication**: It cannot guarantee randomness or unpredictability, unlike cryptographic hash functions (e.g., SHA-256).

##### 4. **Limited Error Detection**
- **Weakness**: CRC32 detects only limited types of errors.
- **Reason**: It is primarily designed to detect single-bit errors or small groups of bit errors. Complex or systematic errors can remain undetected.
- **Implication**: It is insufficient to ensure data integrity in modern communication systems.

##### 5. **Not Universally Suitable**
- **Weakness**: CRC32 is not flexible for varying data sizes.
- **Reason**: It operates on bit or byte levels and does not scale well for very large datasets or different structures.
- **Implication**: For large amounts of data, hash algorithms with better scalability and larger outputs (e.g., SHA-256 with 256 bits) are preferable.

##### 6. **No Unified Implementation**
- **Weakness**: There are multiple variants of CRC32, differing by the polynomial constant used (e.g., CRC32C, CRC32-BZIP2).
- **Reason**: Different applications adopt different standards.
- **Implication**: Hashes may be incompatible across implementations.

#### When is CRC32 Useful?
CRC32 is well-suited for:
- **Checksums**: Verifying data integrity for files or data transmission.
- **Error Detection**: Identifying small, random errors in data blocks.
- **Non-Critical Hashing**: Applications where speed and simplicity are more important than security.