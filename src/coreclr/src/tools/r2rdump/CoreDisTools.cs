﻿// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using ILCompiler.Reflection.ReadyToRun;
using System;
using System.Collections.Generic;
using System.Reflection.PortableExecutable;
using System.Runtime.InteropServices;
using System.Text;

namespace R2RDump
{
    public class CoreDisTools
    {
        private const string _dll = "coredistools";

        public enum TargetArch
        {
            Target_Host, // Target is the same as host architecture
            Target_X86,
            Target_X64,
            Target_Thumb,
            Target_Arm64
        };

        [DllImport(_dll)]
        public static extern IntPtr InitBufferedDisasm(TargetArch Target);

        [DllImport(_dll)]
        public static extern void DumpCodeBlock(IntPtr Disasm, ulong Address, IntPtr Bytes, int Size);

        [DllImport(_dll)]
        [return: MarshalAs(UnmanagedType.I4)]
        public static extern int DumpInstruction(IntPtr Disasm, ulong Address, IntPtr Bytes, int Size);

        [DllImport(_dll)]
        public static extern IntPtr GetOutputBuffer();

        [DllImport(_dll)]
        public static extern void ClearOutputBuffer();

        [DllImport(_dll)]
        public static extern void FinishDisasm(IntPtr Disasm);

        public unsafe static int GetInstruction(IntPtr Disasm, RuntimeFunction rtf, int imageOffset, int rtfOffset, byte[] image, out string instr)
        {
            int instrSize = 1;
            fixed (byte* p = image)
            {
                IntPtr ptr = (IntPtr)(p + imageOffset + rtfOffset);
                instrSize = DumpInstruction(Disasm, (ulong)(rtf.StartAddress + rtfOffset), ptr, rtf.Size);
            }
            IntPtr pBuffer = GetOutputBuffer();
            instr = Marshal.PtrToStringAnsi(pBuffer);
            return instrSize;
        }

        public static IntPtr GetDisasm(Machine machine)
        {
            TargetArch target = TargetArch.Target_Host;
            switch (machine)
            {
                case Machine.Amd64:
                    target = TargetArch.Target_X64;
                    break;
                case Machine.I386:
                    target = TargetArch.Target_X86;
                    break;
                case Machine.Arm64:
                    target = TargetArch.Target_Arm64;
                    break;
                case Machine.ArmThumb2:
                    target = TargetArch.Target_Thumb;
                    break;
                default:
                    R2RDump.WriteWarning($"{machine} not supported on CoreDisTools");
                    return IntPtr.Zero;
            }
            return InitBufferedDisasm(target);
        }
    }

    /// <summary>
    /// Helper class for converting machine instructions to textual representation.
    /// </summary>
    public class Disassembler : IDisposable
    {
        /// <summary>
        /// R2R reader is used to access architecture info, the PE image data and symbol table.
        /// </summary>
        private readonly ReadyToRunReader _reader;

        /// <summary>
        /// Dump options
        /// </summary>
        private readonly DumpOptions _options;

        /// <summary>
        /// COM interface to the native disassembler in the CoreDisTools.dll library.
        /// </summary>
        private readonly IntPtr _disasm;

        /// <summary>
        /// Store the R2R reader and construct the disassembler for the appropriate architecture.
        /// </summary>
        /// <param name="reader"></param>
        public Disassembler(ReadyToRunReader reader, DumpOptions options)
        {
            _reader = reader;
            _options = options;
            _disasm = CoreDisTools.GetDisasm(_reader.Machine);
        }

        /// <summary>
        /// Shut down the native disassembler interface.
        /// </summary>
        public void Dispose()
        {
            if (_disasm != IntPtr.Zero)
            {
                CoreDisTools.FinishDisasm(_disasm);
            }
        }

        /// <summary>
        /// Parse a single instruction and return the RVA of the next instruction.
        /// </summary>
        /// <param name="rtf">Runtime function to parse</param>
        /// <param name="imageOffset">Offset within the PE image byte array</param>
        /// <param name="rtfOffset">Instruction offset within the runtime function</param>
        /// <param name="instruction">Output text representation of the instruction</param>
        /// <returns>Instruction size in bytes - i.o.w. the next instruction starts at rtfOffset + (the return value)</returns>
        public int GetInstruction(RuntimeFunction rtf, int imageOffset, int rtfOffset, out string instruction)
        {
            if (_disasm == IntPtr.Zero)
            {
                instruction = "";
                return rtf.Size;
            }

            int instrSize = CoreDisTools.GetInstruction(_disasm, rtf, imageOffset, rtfOffset, _reader.Image, out instruction);
            CoreDisTools.ClearOutputBuffer();

            instruction = instruction.Replace('\t', ' ');

            if (_options.Naked)
            {
                StringBuilder nakedInstruction = new StringBuilder();
                foreach (string line in instruction.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    int colon = line.IndexOf(':');
                    if (colon >= 0)
                    {
                        colon += 2;
                        while (colon + 3 <= line.Length &&
                            IsXDigit(line[colon]) &&
                            IsXDigit(line[colon + 1]) &&
                            line[colon + 2] == ' ')
                        {
                            colon += 3;
                        }

                        if (!_options.HideOffsets)
                        {
                            nakedInstruction.Append($"{(rtfOffset + rtf.CodeOffset),8:x4}:");
                            nakedInstruction.Append("  ");
                        }
                        else
                        {
                            nakedInstruction.Append("    ");
                        }
                        nakedInstruction.Append(line.Substring(colon).TrimStart());
                        nakedInstruction.Append('\n');
                    }
                    else
                    {
                        nakedInstruction.Append(' ', 7);
                        nakedInstruction.Append(line.TrimStart());
                        nakedInstruction.Append('\n');
                    }
                }
                instruction = nakedInstruction.ToString();
            }

            switch (_reader.Machine)
            {
                case Machine.Amd64:
                    ProbeX64Quirks(rtf, imageOffset, rtfOffset, instrSize, ref instruction);
                    break;

                case Machine.I386:
                    ProbeX86Quirks(rtf, imageOffset, rtfOffset, instrSize, ref instruction);
                    break;

                case Machine.ArmThumb2:
                case Machine.Thumb:
                    break;

                case Machine.Arm64:
                    break;

                default:
                    throw new NotImplementedException();
            }

            instruction = instruction.Replace("\n", Environment.NewLine);
            return instrSize;
        }

        private static bool IsXDigit(char c)
        {
            return Char.IsDigit(c) || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f');
        }

        const string RelIPTag = "[rip ";

        /// <summary>
        /// Translate RIP-relative offsets to RVA's and convert cell addresses to symbol names
        /// </summary>
        /// <param name="rtf">Runtime function</param>
        /// <param name="imageOffset">Offset within the image byte array</param>
        /// <param name="rtfOffset">Offset within the runtime function</param>
        /// <param name="instrSize">Instruction size</param>
        /// <param name="instruction">Textual representation of the instruction</param>
        private void ProbeX64Quirks(RuntimeFunction rtf, int imageOffset, int rtfOffset, int instrSize, ref string instruction)
        {
            int leftBracket;
            int rightBracketPlusOne;
            int displacement;
            if (TryParseRipRelative(instruction, out leftBracket, out rightBracketPlusOne, out displacement))
            {
                int target = rtf.StartAddress + rtfOffset + instrSize + displacement;
                int newline = instruction.LastIndexOf('\n');
                StringBuilder translated = new StringBuilder();
                translated.Append(instruction, 0, leftBracket);
                if (_options.Naked)
                {
                    String targetName;
                    if (_reader.ImportCellNames.TryGetValue(target, out targetName))
                    {
                        translated.AppendFormat("[{0}]", targetName);
                    }
                    else
                    {
                        translated.AppendFormat("[0x{0:x4}]", target);
                    }
                }
                else
                {
                    translated.AppendFormat("[0x{0:x4}]", target);

                    AppendImportCellName(translated, target);
                }

                translated.Append(instruction, rightBracketPlusOne, newline - rightBracketPlusOne);

                translated.Append(instruction, newline, instruction.Length - newline);
                instruction = translated.ToString();
            }
            else
            {
                ProbeCommonIntelQuirks(rtf, imageOffset, rtfOffset, instrSize, ref instruction);
            }
        }

        /// <summary>
        /// X86 disassembler has a bug in decoding absolute indirections, mistaking them for RIP-relative indirections
        /// </summary>
        /// <param name="rtf">Runtime function</param>
        /// <param name="imageOffset">Offset within the image byte array</param>
        /// <param name="rtfOffset">Offset within the runtime function</param>
        /// <param name="instrSize">Instruction size</param>
        /// <param name="instruction">Textual representation of the instruction</param>
        private void ProbeX86Quirks(RuntimeFunction rtf, int imageOffset, int rtfOffset, int instrSize, ref string instruction)
        {
            int leftBracket;
            int rightBracketPlusOne;
            int absoluteAddress;
            if (TryParseRipRelative(instruction, out leftBracket, out rightBracketPlusOne, out absoluteAddress))
            {
                int target = absoluteAddress - (int)_reader.PEReader.PEHeaders.PEHeader.ImageBase;

                StringBuilder translated = new StringBuilder();
                translated.Append(instruction, 0, leftBracket);
                if (_options.Naked)
                {
                    String targetName;
                    if (_reader.ImportCellNames.TryGetValue(target, out targetName))
                    {
                        translated.AppendFormat("[{0}]", targetName);
                    }
                    else
                    {
                        translated.AppendFormat("[0x{0:x4}]", target);
                    }
                }
                else
                {
                    translated.AppendFormat("[0x{0:x4}]", target);

                    AppendImportCellName(translated, target);
                }

                translated.Append(instruction, rightBracketPlusOne, instruction.Length - rightBracketPlusOne);
                instruction = translated.ToString();
            }
            else
            {
                ProbeCommonIntelQuirks(rtf, imageOffset, rtfOffset, instrSize, ref instruction);
            }
        }

        /// <summary>
        /// Probe quirks that have the same behavior for X86 and X64.
        /// </summary>
        /// <param name="rtf">Runtime function</param>
        /// <param name="imageOffset">Offset within the image byte array</param>
        /// <param name="rtfOffset">Offset within the runtime function</param>
        /// <param name="instrSize">Instruction size</param>
        /// <param name="instruction">Textual representation of the instruction</param>
        private void ProbeCommonIntelQuirks(RuntimeFunction rtf, int imageOffset, int rtfOffset, int instrSize, ref string instruction)
        {
            int instructionRVA = rtf.StartAddress + rtfOffset;
            int nextInstructionRVA = instructionRVA + instrSize;
            if (instrSize == 2 && IsIntelJumpInstructionWithByteOffset(imageOffset + rtfOffset))
            {
                sbyte offset = (sbyte)_reader.Image[imageOffset + rtfOffset + 1];
                ReplaceRelativeOffset(ref instruction, nextInstructionRVA + offset, rtf);
            }
            else if (instrSize == 5 && IsIntel1ByteJumpInstructionWithIntOffset(imageOffset + rtfOffset))
            {
                int offset = BitConverter.ToInt32(_reader.Image, imageOffset + rtfOffset + 1);
                ReplaceRelativeOffset(ref instruction, nextInstructionRVA + offset, rtf);
            }
            else if (instrSize == 5 && IsIntelCallInstructionWithIntOffset(imageOffset + rtfOffset))
            {
                int offset = BitConverter.ToInt32(_reader.Image, imageOffset + rtfOffset + 1);
                int targetRVA = nextInstructionRVA + offset;
                int targetImageOffset = _reader.GetOffset(targetRVA);
                bool pointsOutsideRuntimeFunction = (targetRVA < rtf.StartAddress || targetRVA >= rtf.StartAddress + rtf.Size);
                if (pointsOutsideRuntimeFunction && IsIntel2ByteIndirectJumpPCRelativeInstruction(targetImageOffset, out int instructionRelativeOffset))
                {
                    int thunkTargetRVA = targetRVA + instructionRelativeOffset;
                    bool haveImportCell = _reader.ImportCellNames.TryGetValue(thunkTargetRVA, out string importCellName);

                    if (_options.Naked && haveImportCell)
                    {
                        ReplaceRelativeOffset(ref instruction, $@"qword ptr [{importCellName}]", rtf);
                    }
                    else
                    {
                        ReplaceRelativeOffset(ref instruction, targetRVA, rtf);
                        if (haveImportCell)
                        {
                            int instructionEnd = instruction.IndexOf('\n');
                            StringBuilder builder = new StringBuilder(instruction, 0, instructionEnd, capacity: 256);
                            AppendComment(builder, @$"JMP [0x{thunkTargetRVA:X4}]: {importCellName}");
                            builder.AppendLine();
                            instruction = builder.ToString();
                        }
                    }
                }
                else if (pointsOutsideRuntimeFunction && IsAnotherRuntimeFunctionWithinMethod(targetRVA, rtf, out int runtimeFunctionIndex))
                {
                    string runtimeFunctionName = string.Format("RUNTIME_FUNCTION[{0}]", runtimeFunctionIndex);

                    if (_options.Naked)
                    {
                        ReplaceRelativeOffset(ref instruction, runtimeFunctionName, rtf);
                    }
                    else
                    {
                        ReplaceRelativeOffset(ref instruction, targetRVA, rtf);
                        int instructionEnd = instruction.IndexOf('\n');
                        StringBuilder builder = new StringBuilder(instruction, 0, instructionEnd, capacity: 256);
                        AppendComment(builder, runtimeFunctionName);
                        builder.AppendLine();
                        instruction = builder.ToString();
                    }
                }
                else
                {
                    ReplaceRelativeOffset(ref instruction, nextInstructionRVA + offset, rtf);
                }
            }
            else if (instrSize == 6 && IsIntel2ByteJumpInstructionWithIntOffset(imageOffset + rtfOffset))
            {
                int offset = BitConverter.ToInt32(_reader.Image, imageOffset + rtfOffset + 2);
                ReplaceRelativeOffset(ref instruction, nextInstructionRVA + offset, rtf);
            }
        }

        /// <summary>
        /// Try to parse the [rip +- displacement] section in a disassembled instruction string.
        /// </summary>
        /// <param name="instruction">Disassembled instruction string</param>
        /// <param name="leftBracket">Index of the left bracket in the instruction</param>
        /// <param name="rightBracketPlusOne">Index of the right bracket in the instruction plus one</param>
        /// <param name="displacement">Value of the IP-relative delta</param>
        /// <returns></returns>
        private bool TryParseRipRelative(string instruction, out int leftBracket, out int rightBracket, out int displacement)
        {
            int relip = instruction.IndexOf(RelIPTag);
            if (relip >= 0 && instruction.Length >= relip + RelIPTag.Length + 3)
            {
                int start = relip;
                relip += RelIPTag.Length;
                char sign = instruction[relip];
                if (sign == '+' || sign == '-' &&
                    instruction[relip + 1] == ' ' &&
                    Char.IsDigit(instruction[relip + 2]))
                {
                    relip += 2;
                    int offset = 0;
                    do
                    {
                        offset = 10 * offset + (int)(instruction[relip] - '0');
                    }
                    while (++relip < instruction.Length && Char.IsDigit(instruction[relip]));
                    if (relip < instruction.Length && instruction[relip] == ']')
                    {
                        relip++;
                        if (sign == '-')
                        {
                            offset = -offset;
                        }
                        leftBracket = start;
                        rightBracket = relip;
                        displacement = offset;
                        return true;
                    }
                }
            }

            leftBracket = 0;
            rightBracket = 0;
            displacement = 0;
            return false;
        }

        /// <summary>
        /// Append import cell name to the constructed instruction string as a comment if available.
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="importCellRva"></param>
        private void AppendImportCellName(StringBuilder builder, int importCellRva)
        {
            String targetName;
            if (_reader.ImportCellNames.TryGetValue(importCellRva, out targetName))
            {
                AppendComment(builder, targetName);
            }
        }

        /// <summary>
        /// Append a given comment to the string builder.
        /// </summary>
        /// <param name="builder">String builder to append comment to</param>
        /// <param name="comment">Comment to append</param>
        private void AppendComment(StringBuilder builder, string comment)
        {
            int fill = 61 - builder.Length;
            if (fill > 0)
            {
                builder.Append(' ', fill);
            }
            builder.Append(" // ");
            builder.Append(comment);
        }

        /// <summary>
        /// Replace relative offset in the disassembled instruction with the true target RVA.
        /// </summary>
        /// <param name="instruction">Disassembled instruction to modify</param>
        /// <param name="target">Target string to replace offset with</param>
        /// <param name="rtf">Runtime function being disassembled</param>
        private void ReplaceRelativeOffset(ref string instruction, int target, RuntimeFunction rtf)
        {
            int outputOffset = target;
            if (_options.Naked)
            {
                outputOffset -= rtf.StartAddress;
            }
            ReplaceRelativeOffset(ref instruction, string.Format("0x{0:X4}", outputOffset), rtf);
        }

        /// <summary>
        /// Replace relative offset in the disassembled instruction with an arbitrary string.
        /// </summary>
        /// <param name="instruction">Disassembled instruction to modify</param>
        /// <param name="replacementString">String to replace offset with</param>
        /// <param name="rtf">Runtime function being disassembled</param>
        private void ReplaceRelativeOffset(ref string instruction, string replacementString, RuntimeFunction rtf)
        {
            int numberEnd = instruction.IndexOf('\n');
            int number = numberEnd;
            while (number > 0)
            {
                char c = instruction[number - 1];
                if (c >= ' ' && !Char.IsDigit(c) && c != '-')
                {
                    break;
                }
                number--;
            }

            StringBuilder translated = new StringBuilder();
            translated.Append(instruction, 0, number);
            translated.Append(replacementString);
            translated.Append(instruction, numberEnd, instruction.Length - numberEnd);
            instruction = translated.ToString();
        }

        /// <summary>
        /// Returns true when this is one of the x86 / amd64 opcodes used for branch instructions
        /// with single-byte offset.
        /// </summary>
        /// <param name="imageOffset">Offset within the PE image byte array</param>
        private bool IsIntelJumpInstructionWithByteOffset(int imageOffset)
        {
            byte opCode = _reader.Image[imageOffset];
            return
                (opCode >= 0x70 && opCode <= 0x7F) // short conditional jumps
                || opCode == 0xE3 // JCXZ
                || opCode == 0xEB // JMP
                ;
        }

        /// <summary>
        /// Returns true for the call relative opcode with signed 4-byte offset.
        /// </summary>
        /// <param name="imageOffset">Offset within the PE image byte array</param>
        private bool IsIntelCallInstructionWithIntOffset(int imageOffset)
        {
            return _reader.Image[imageOffset] == 0xE8; // CALL rel32
        }

        /// <summary>
        /// Returns true when this is one of the x86 / amd64 near jump / call opcodes
        /// with signed 4-byte offset.
        /// </summary>
        /// <param name="imageOffset">Offset within the PE image byte array</param>
        private bool IsIntel1ByteJumpInstructionWithIntOffset(int imageOffset)
        {
            return _reader.Image[imageOffset] == 0xE9; // JMP rel32
        }

        /// <summary>
        /// Returns true when this is one of the x86 / amd64 conditional near jump
        /// opcodes with signed 4-byte offset.
        /// </summary>
        /// <param name="imageOffset">Offset within the PE image byte array</param>
        private bool IsIntel2ByteJumpInstructionWithIntOffset(int imageOffset)
        {
            byte opCode1 = _reader.Image[imageOffset];
            byte opCode2 = _reader.Image[imageOffset + 1];
            return opCode1 == 0x0F &&
                (opCode2 >= 0x80 && opCode2 <= 0x8F); // near conditional jumps
        }

        /// <summary>
        /// Returns true when this is the 2-byte instruction for indirect jump
        /// with RIP-relative offset.
        /// </summary>
        /// <param name="imageOffset">Offset within the PE image byte array</param>
        /// <returns></returns>
        private bool IsIntel2ByteIndirectJumpPCRelativeInstruction(int imageOffset, out int instructionRelativeOffset)
        {
            byte opCode1 = _reader.Image[imageOffset + 0];
            byte opCode2 = _reader.Image[imageOffset + 1];
            int offsetDelta = 6;

            if (opCode1 == 0x48 && opCode2 == 0x8B && _reader.Image[imageOffset + 2] == 0x15) // MOV RDX, [R2R module]
            {
                imageOffset += 7;
                offsetDelta += 7;
                opCode1 = _reader.Image[imageOffset + 0];
                opCode2 = _reader.Image[imageOffset + 1];
            }

            if (opCode1 == 0xFF && opCode2 == 0x25)
            {
                // JMP [RIP + rel32]
                instructionRelativeOffset = offsetDelta + BitConverter.ToInt32(_reader.Image, imageOffset + 2);
                return true;
            }

            instructionRelativeOffset = 0;
            return false;
        }

        /// <summary>
        /// Check whether a given target RVA corresponds to another runtime function within the same method.
        /// </summary>
        /// <param name="rva">Target RVA to analyze</param>
        /// <param name="rtf">Runtime function being disassembled</param>
        /// <param name="runtimeFunctionIndex">Output runtime function index if found, -1 otherwise</param>
        /// <returns>true if target runtime function has been found, false otherwise</returns>
        private bool IsAnotherRuntimeFunctionWithinMethod(int rva, RuntimeFunction rtf, out int runtimeFunctionIndex)
        {
            for (int rtfIndex = 0; rtfIndex < rtf.Method.RuntimeFunctions.Count; rtfIndex++)
            {
                if (rva == rtf.Method.RuntimeFunctions[rtfIndex].StartAddress)
                {
                    runtimeFunctionIndex = rtfIndex;
                    return true;
                }
            }

            runtimeFunctionIndex = -1;
            return false;
        }
    }
}
