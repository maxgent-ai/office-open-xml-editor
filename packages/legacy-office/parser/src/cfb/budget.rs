//! CFB materialization policy, separate from MS-CFB validity. Reuse the legacy
//! readers' 256 MiB retention and 2,000,000 work scales for this shared layer.
//! The borrowed source length and stream declarations never consume this budget.
//! Vector capacities, retained names, and hash-slot allowances consume retention;
//! decoded FAT/DIFAT entries and visited chain/directory links consume work. Charges are cumulative for one CompoundFile, including
//! temporary buffers and repeated reads; releasing a buffer does not replenish
//! work or allow repeated materialization to evade the ceiling. Allocation or
//! quota failure poisons subsequent reads instead of returning partial streams.
use std::cell::Cell;
use std::collections::{HashMap, HashSet};
use std::hash::Hash;
use std::mem::size_of;

const MAX_RETAINED_BYTES: usize = 256 * 1024 * 1024;
const MAX_WORK: usize = 2_000_000;

pub(super) struct Budget {
    bytes: Cell<usize>,
    work: Cell<usize>,
    failed: Cell<bool>,
}

impl Default for Budget {
    fn default() -> Self {
        Self {
            bytes: Cell::new(MAX_RETAINED_BYTES),
            work: Cell::new(MAX_WORK),
            failed: Cell::new(false),
        }
    }
}

impl Budget {
    pub(super) fn resource(&self, message: &str) -> String {
        self.failed.set(true);
        crate::rejection::resource(message)
    }

    fn fail(&self) -> String {
        self.resource("CFB retention/work budget or allocation exhausted")
    }

    #[cfg(test)]
    pub(super) fn with_limits(bytes: usize, work: usize) -> Self {
        Self {
            bytes: Cell::new(bytes),
            work: Cell::new(work),
            failed: Cell::new(false),
        }
    }

    pub(super) fn check(&self) -> Result<(), String> {
        if self.failed.get() {
            Err(self.fail())
        } else {
            Ok(())
        }
    }

    fn charge(&self, counter: &Cell<usize>, amount: usize) -> Result<(), String> {
        self.check()?;
        counter.set(
            counter
                .get()
                .checked_sub(amount)
                .ok_or_else(|| self.fail())?,
        );
        Ok(())
    }

    pub(super) fn step(&self) -> Result<(), String> {
        self.charge(&self.work, 1)
    }

    /// Grow from actual appended elements, with geometric amortization and no
    /// declaration-sized reservation. Account capacity (including spare space)
    /// before retaining more data. Exact reservation avoids an uncharged extra
    /// growth chosen by Vec; any allocator-reported excess is also charged.
    pub(super) fn reserve<T>(&self, vec: &mut Vec<T>, additional: usize) -> Result<(), String> {
        self.check()?;
        let required = vec
            .len()
            .checked_add(additional)
            .ok_or_else(|| self.fail())?;
        if required <= vec.capacity() {
            return Ok(());
        }
        let unit = size_of::<T>().max(1);
        let available = self.bytes.get() / unit;
        let needed = required - vec.capacity();
        if needed > available {
            return Err(self.fail());
        }
        let growth = vec.capacity().max(needed).min(available);
        let before = vec.capacity();
        vec.try_reserve_exact(before + growth - vec.len())
            .map_err(|_| self.fail())?;
        self.charge(
            &self.bytes,
            (vec.capacity() - before)
                .checked_mul(unit)
                .ok_or_else(|| self.fail())?,
        )
    }

    pub(super) fn push<T>(&self, vec: &mut Vec<T>, value: T) -> Result<(), String> {
        self.reserve(vec, 1)?;
        vec.push(value);
        Ok(())
    }

    pub(super) fn append(&self, vec: &mut Vec<u8>, bytes: &[u8]) -> Result<(), String> {
        self.reserve(vec, bytes.len())?;
        vec.extend_from_slice(bytes);
        Ok(())
    }

    /// Hash tables reserve fallibly per inserted item. Charge a conservative
    /// slot allowance for control bytes, load factor, and geometric spare slots.
    /// Four slots per item cover the small-table minimum and doubling, with a
    /// word per slot for control/alignment overhead. This is resource policy,
    /// not an MS-CFB limit or a reservation from declared sector counts.
    pub(super) fn insert<T: Eq + Hash>(
        &self,
        set: &mut HashSet<T>,
        value: T,
    ) -> Result<bool, String> {
        self.check()?;
        if set.contains(&value) {
            return Ok(false);
        }
        self.charge(&self.bytes, 4 * (size_of::<T>() + size_of::<usize>()))?;
        set.try_reserve(1).map_err(|_| self.fail())?;
        Ok(set.insert(value))
    }

    pub(super) fn index(
        &self,
        map: &mut HashMap<String, usize>,
        name: String,
        id: usize,
    ) -> Result<bool, String> {
        self.charge(
            &self.bytes,
            4 * (size_of::<(String, usize)>() + size_of::<usize>()),
        )?;
        map.try_reserve(1).map_err(|_| self.fail())?;
        Ok(map.insert(name, id).is_some())
    }

    pub(super) fn name(
        &self,
        chars: impl Iterator<Item = char>,
        capacity: usize,
    ) -> Result<String, String> {
        self.charge(&self.bytes, capacity)?;
        let mut name = String::new();
        name.try_reserve_exact(capacity).map_err(|_| self.fail())?;
        if name.capacity() > capacity {
            self.charge(&self.bytes, name.capacity() - capacity)?;
        }
        // Callers supply an upper bound derived from the actual bounded name,
        // never a stream declaration. This loop cannot trigger further growth.
        for c in chars {
            name.push(c);
        }
        Ok(name)
    }
}
